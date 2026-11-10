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
            String contentRange, int bodyLength, String prefer, String range) {
    }

    private final HttpServer server;
    private final Map<String, Stub> stubs = new ConcurrentHashMap<>();
    private final Map<String, Queue<Stub>> onceStubs = new ConcurrentHashMap<>();
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
        requests.add(new RecordedRequest(exchange.getRequestMethod(), path,
                query == null ? "" : query,
                exchange.getRequestHeaders().getFirst("authorization"),
                exchange.getRequestHeaders().getFirst("content-range"),
                body.length,
                exchange.getRequestHeaders().getFirst("prefer"),
                exchange.getRequestHeaders().getFirst("range")));

        Stub stub = null;
        Queue<Stub> once = onceStubs.get(path);
        if (once != null) {
            stub = once.poll();
        }
        if (stub == null && query != null) {
            stub = stubs.get(path + "?" + query);
        }
        if (stub == null) {
            stub = stubs.get(path);
        }
        if (stub == null && "PUT".equals(exchange.getRequestMethod())) {
            stub = Stub.json("{\"id\":\"uploaded\",\"name\":\"file\"}");
        }
        if (stub == null) {
            stub = Stub.json("{\"error\":{\"code\":\"notFound\",\"message\":\"not stubbed: "
                    + path + "\"}}");
            stub = new Stub(404, stub.body(), stub.headers());
        }

        stub.headers().forEach((name, value) -> exchange.getResponseHeaders().add(name, value));
        exchange.sendResponseHeaders(stub.status(), stub.body().length);
        try (OutputStream out = exchange.getResponseBody()) {
            out.write(stub.body());
        }
    }

    @Override
    public void close() {
        server.stop(0);
    }
}
