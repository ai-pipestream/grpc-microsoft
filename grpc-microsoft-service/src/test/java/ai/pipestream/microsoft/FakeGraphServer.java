package ai.pipestream.microsoft;

import com.sun.net.httpserver.HttpServer;

import java.io.IOException;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * In-process fake of Microsoft Graph over {@link HttpServer}. Stubs are
 * keyed by request path (optionally {@code path?query}).
 */
final class FakeGraphServer implements AutoCloseable {

    record Stub(int status, byte[] body, Map<String, String> headers) {
        static Stub json(String body) {
            return new Stub(200, body.getBytes(StandardCharsets.UTF_8),
                    Map.of("content-type", "application/json"));
        }

        static Stub bytes(byte[] body, String contentType) {
            return new Stub(200, body, Map.of("content-type", contentType));
        }
    }

    record RecordedRequest(String method, String path, String query, String authorization,
            String contentRange, int bodyLength, String prefer, String range, String ifMatch,
            String body) {
    }

    /** Answers a request from its recorded form. */
    interface Responder {
        Stub respond(RecordedRequest request);
    }

    private final HttpServer server;
    private final Map<String, Stub> stubs = new ConcurrentHashMap<>();
    private final Map<String, Queue<Stub>> onceStubs = new ConcurrentHashMap<>();
    private final Map<String, Responder> responders = new ConcurrentHashMap<>();
    private final Map<String, java.util.concurrent.atomic.AtomicReference<byte[]>> lastBodies =
            new ConcurrentHashMap<>();
    private final List<RecordedRequest> requests = new CopyOnWriteArrayList<>();

    private FakeGraphServer(HttpServer server) {
        this.server = server;
    }

    static FakeGraphServer start() throws IOException {
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        FakeGraphServer fake = new FakeGraphServer(server);
        server.createContext("/", exchange -> {
            try {
                fake.handle(exchange);
            } finally {
                exchange.close();
            }
        });
        server.start();
        return fake;
    }

    String baseUrl() {
        return "http://127.0.0.1:" + server.getAddress().getPort();
    }

    void stub(String pathOrPathQuery, String json) {
        stubs.put(pathOrPathQuery, Stub.json(json));
    }

    void stub(String pathOrPathQuery, Stub stub) {
        stubs.put(pathOrPathQuery, stub);
    }

    void stubOnce(String path, Stub stub) {
        onceStubs.computeIfAbsent(path, k -> new ConcurrentLinkedQueue<>()).add(stub);
    }

    /** A stub for one HTTP method only; wins over a method-agnostic stub. */
    void stub(String method, String path, Stub stub) {
        stubs.put(method + " " + path, stub);
    }

    /** Answers every request to {@code path} with {@code responder}. */
    void respond(String path, Responder responder) {
        responders.put(path, responder);
    }

    /** A JSON stub with a status and headers. */
    static Stub json(int status, String body, Map<String, String> headers) {
        java.util.Map<String, String> all = new java.util.HashMap<>(headers);
        all.put("content-type", "application/json");
        return new Stub(status, body.getBytes(StandardCharsets.UTF_8), all);
    }

    /** A bodiless stub. */
    static Stub empty(int status, Map<String, String> headers) {
        return new Stub(status, new byte[0], headers);
    }

    /**
     * An upload session at {@code path}: ranged PUTs must arrive in order, a GET
     * reports the next expected offset, and the PUT that completes the file
     * answers {@code itemJson}. Returns the bytes received so far.
     */
    java.io.ByteArrayOutputStream uploadSession(String path, String itemJson) {
        java.io.ByteArrayOutputStream stored = new java.io.ByteArrayOutputStream();
        java.util.concurrent.atomic.AtomicReference<byte[]> last =
                new java.util.concurrent.atomic.AtomicReference<>();
        lastBodies.put(path, last);
        respond(path, request -> {
            synchronized (stored) {
                String next = "{\"nextExpectedRanges\":[\"" + stored.size() + "-\"]}";
                if ("DELETE".equals(request.method())) {
                    return empty(204, Map.of());
                }
                if (!"PUT".equals(request.method())) {
                    return Stub.json(next);
                }
                String range = request.contentRange();
                long start = Long.parseLong(range.substring("bytes ".length(), range.indexOf('-')));
                long end = Long.parseLong(range.substring(range.indexOf('-') + 1, range.indexOf('/')));
                long total = Long.parseLong(range.substring(range.indexOf('/') + 1));
                if (start != stored.size() || end - start + 1 != request.bodyLength()) {
                    return json(416, next, Map.of());
                }
                stored.writeBytes(last.get());
                if (stored.size() == total) {
                    return json(201, itemJson, Map.of());
                }
                return json(202, "{\"nextExpectedRanges\":[\"" + stored.size() + "-\"]}",
                        Map.of());
            }
        });
        return stored;
    }

    /**
     * Serves {@code POST /$batch} by answering each sub-request from the stubs,
     * as if it had been sent on its own. Sub-requests are recorded too.
     */
    void enableBatch() {
        respond("/$batch", request -> {
            try {
                com.fasterxml.jackson.databind.ObjectMapper json =
                        new com.fasterxml.jackson.databind.ObjectMapper();
                com.fasterxml.jackson.databind.node.ObjectNode answer = json.createObjectNode();
                com.fasterxml.jackson.databind.node.ArrayNode responses =
                        answer.putArray("responses");
                for (com.fasterxml.jackson.databind.JsonNode sub
                        : json.readTree(request.body()).path("requests")) {
                    String url = sub.path("url").asText();
                    int mark = url.indexOf('?');
                    String path = java.net.URLDecoder.decode(mark < 0 ? url : url.substring(0, mark),
                            StandardCharsets.UTF_8);
                    String query = mark < 0 ? "" : url.substring(mark + 1);
                    String body = sub.has("body") ? sub.path("body").toString() : "";
                    RecordedRequest recorded = new RecordedRequest(sub.path("method").asText(),
                            path, query, request.authorization(), null, body.length(), null, null,
                            sub.path("headers").path("If-Match").asText(null), body);
                    requests.add(recorded);
                    Stub stub = resolve(recorded);
                    com.fasterxml.jackson.databind.node.ObjectNode response = responses.addObject()
                            .put("id", sub.path("id").asText())
                            .put("status", stub.status());
                    com.fasterxml.jackson.databind.node.ObjectNode headers =
                            response.putObject("headers");
                    stub.headers().forEach(headers::put);
                    if (stub.body().length > 0) {
                        response.set("body", json.readTree(stub.body()));
                    }
                }
                return Stub.json(answer.toString());
            } catch (IOException e) {
                throw new java.io.UncheckedIOException(e);
            }
        });
    }

    List<RecordedRequest> requests() {
        return List.copyOf(requests);
    }

    /** Requests whose path is exactly {@code path}, in order. */
    List<RecordedRequest> requestsTo(String path) {
        return requests.stream().filter(r -> r.path().equals(path)).toList();
    }

    /** The single request to {@code path}; fails when there were none or several. */
    RecordedRequest only(String path) {
        List<RecordedRequest> matching = requestsTo(path);
        if (matching.size() != 1) {
            throw new AssertionError("expected one request to " + path + " but saw " + matching
                    + " among " + requests);
        }
        return matching.get(0);
    }

    /** A JSON error stub. */
    static Stub error(int status, String code, Map<String, String> headers) {
        java.util.Map<String, String> all = new java.util.HashMap<>(headers);
        all.put("content-type", "application/json");
        return new Stub(status, ("{\"error\":{\"code\":\"" + code + "\",\"message\":\"" + code
                + "\"}}").getBytes(StandardCharsets.UTF_8), all);
    }

    private void handle(com.sun.net.httpserver.HttpExchange exchange) throws IOException {
        String path = exchange.getRequestURI().getPath();
        String query = exchange.getRequestURI().getRawQuery();
        byte[] body = exchange.getRequestBody().readAllBytes();
        String contentType = exchange.getRequestHeaders().getFirst("content-type");
        boolean text = contentType != null && contentType.startsWith("application/json");
        RecordedRequest recorded = new RecordedRequest(exchange.getRequestMethod(), path,
                query == null ? "" : query,
                exchange.getRequestHeaders().getFirst("authorization"),
                exchange.getRequestHeaders().getFirst("content-range"),
                body.length,
                exchange.getRequestHeaders().getFirst("prefer"),
                exchange.getRequestHeaders().getFirst("range"),
                exchange.getRequestHeaders().getFirst("if-match"),
                text ? new String(body, StandardCharsets.UTF_8) : "");
        requests.add(recorded);
        java.util.concurrent.atomic.AtomicReference<byte[]> last = lastBodies.get(path);
        if (last != null) {
            last.set(body);
        }

        Stub stub = resolve(recorded);
        stub.headers().forEach((name, value) -> exchange.getResponseHeaders().add(name, value));
        exchange.sendResponseHeaders(stub.status(), stub.body().length == 0 ? -1 : stub.body().length);
        if (stub.body().length > 0) {
            try (OutputStream out = exchange.getResponseBody()) {
                out.write(stub.body());
            }
        }
    }

    private Stub resolve(RecordedRequest request) {
        String path = request.path();
        Stub stub = null;
        Queue<Stub> once = onceStubs.get(path);
        if (once != null) {
            stub = once.poll();
        }
        Responder responder = responders.get(path);
        if (stub == null && responder != null) {
            stub = responder.respond(request);
        }
        if (stub == null) {
            stub = stubs.get(request.method() + " " + path);
        }
        if (stub == null && !request.query().isEmpty()) {
            stub = stubs.get(path + "?" + request.query());
        }
        if (stub == null) {
            stub = stubs.get(path);
        }
        if (stub == null && "PUT".equals(request.method())) {
            stub = Stub.json("{\"id\":\"uploaded\",\"name\":\"file\"}");
        }
        if (stub == null) {
            stub = Stub.json("{\"error\":{\"code\":\"notFound\",\"message\":\"not stubbed: "
                    + path + "\"}}");
            stub = new Stub(404, stub.body(), stub.headers());
        }
        return stub;
    }

    @Override
    public void close() {
        server.stop(0);
    }
}
