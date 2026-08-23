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

    record RecordedRequest(String method, String path, String query, String authorization) {
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

    private void handle(com.sun.net.httpserver.HttpExchange exchange) throws IOException {
        String path = exchange.getRequestURI().getPath();
        String query = exchange.getRequestURI().getRawQuery();
        requests.add(new RecordedRequest(exchange.getRequestMethod(), path,
                query == null ? "" : query,
                exchange.getRequestHeaders().getFirst("authorization")));

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
