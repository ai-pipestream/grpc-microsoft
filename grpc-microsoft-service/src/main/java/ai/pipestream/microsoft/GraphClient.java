package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import ai.pipestream.microsoft.v1.AuthMode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpHeaders;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.Supplier;

/**
 * A thin authorized client for the Microsoft Graph REST API: JSON in, JSON out, bearer token
 * from a supplier (so token refresh lives with {@link GraphAuth}, not here), and Graph's
 * throttling contract honored — 429/503 responses are retried after the server-directed
 * {@code Retry-After}. No Microsoft SDK: the surface this toolkit needs is a handful of
 * routes, and the whole client stays inspectable.
 */
public final class GraphClient {

    /** A Graph error, with the service's own code and message (client-relevant by design). */
    public static final class GraphApiException extends IOException {

        private static final long serialVersionUID = 1L;

        /** {@link #status()} for a failure Graph never expressed as an HTTP status. */
        public static final int NO_HTTP_STATUS = 0;

        /** HTTP status of the failing response, or {@link #NO_HTTP_STATUS}. */
        private final int status;
        /** Graph error {@code code} field, or empty. */
        private final String code;
        /** Raw error response body. */
        private final String body;
        /** {@code Location} header of the failing response, or empty. */
        private final String location;
        /** {@code Retry-After} seconds of the failing response, or -1. */
        private final long retryAfterSeconds;

        GraphApiException(int status, String code, String message) {
            this(status, code, message, "");
        }

        GraphApiException(int status, String code, String message, String body) {
            this(status, code, message, body, "");
        }

        GraphApiException(int status, String code, String message, String body,
                String location) {
            this(status, code, message, body, location, -1L);
        }

        GraphApiException(int status, String code, String message, String body,
                String location, long retryAfterSeconds) {
            super("Graph " + (status == NO_HTTP_STATUS ? "operation" : Integer.toString(status))
                    + (code.isEmpty() ? "" : " (" + code + ")") + ": " + message);
            this.status = status;
            this.code = code;
            this.body = body;
            this.location = location == null ? "" : location;
            this.retryAfterSeconds = retryAfterSeconds;
        }

        /**
         * The {@code Retry-After} header of the failing response, in seconds.
         *
         * @return the delay; -1 when absent
         */
        public long retryAfterSeconds() {
            return retryAfterSeconds;
        }

        /**
         * The {@code Location} header of the failing response. Graph sets it on a
         * {@code 410 Gone} delta response to the URL of a fresh enumeration.
         *
         * @return the header value; empty when absent
         */
        public String location() {
            return location;
        }

        /**
         * The HTTP status of the failing Graph response, or {@link #NO_HTTP_STATUS} when the
         * failure was not an HTTP one — an async operation that reported failure, or one this
         * client stopped waiting on. Both arrive over a 200 poll, so reporting 200 would name
         * a success.
         *
         * @return the HTTP status, or {@link #NO_HTTP_STATUS}
         */
        public int status() {
            return status;
        }

        /**
         * Graph's error {@code code} field, or empty when the body had none.
         *
         * @return the Graph error code
         */
        public String code() {
            return code;
        }

        /**
         * Graph's full error response body; {@code innerError} often names the real cause.
         *
         * @return the raw body; empty when none was captured
         */
        public String body() {
            return body;
        }
    }

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final int MAX_THROTTLE_RETRIES = 3;

    /** Most redirects {@link #openContent} follows before giving up. */
    private static final int MAX_REDIRECTS = 5;

    private final String baseUrl;
    private final Supplier<String> bearer;
    private final HttpClient http;
    private final AuthMode authMode;

    /**
     * Builds a client against the public Graph v1.0 endpoint.
     *
     * @param bearer supplies a current access token for each request
     */
    public GraphClient(Supplier<String> bearer) {
        this("https://graph.microsoft.com/v1.0", bearer);
    }

    /**
     * Builds a client against {@code baseUrl} (tests and national-cloud
     * endpoints).
     *
     * @param baseUrl Graph root URL, no trailing slash required
     * @param bearer supplies a current access token for each request
     */
    public GraphClient(String baseUrl, Supplier<String> bearer) {
        this(baseUrl, bearer, AuthMode.AUTH_MODE_UNSPECIFIED);
    }

    /**
     * Builds a client against {@code baseUrl} whose tokens are known to come from
     * {@code authMode} (client credentials = application, device code = delegated).
     *
     * @param baseUrl Graph root URL, no trailing slash required
     * @param bearer supplies a current access token for each request
     * @param authMode how the tokens are issued; {@code AUTH_MODE_UNSPECIFIED} lets
     *        {@link #authMode()} inspect the token's claims instead
     */
    public GraphClient(String baseUrl, Supplier<String> bearer, AuthMode authMode) {
        this.baseUrl = GraphLinks.stripTrailingSlashes(Objects.requireNonNull(baseUrl, "baseUrl"));
        this.bearer = Objects.requireNonNull(bearer, "bearer");
        this.http = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(20)).build();
        this.authMode = authMode == null ? AuthMode.AUTH_MODE_UNSPECIFIED : authMode;
    }

    /**
     * The Graph root URL every relative path is resolved against.
     *
     * @return the base URL without a trailing slash
     */
    public String baseUrl() {
        return baseUrl;
    }

    /**
     * Which auth mode the held token has: the mode given at construction, or else
     * the one the current access token's claims declare ({@code idtyp=app} or a
     * {@code roles} claim without {@code scp} = application; {@code scp} =
     * delegated).
     *
     * @return the auth mode; {@code AUTH_MODE_UNSPECIFIED} when the token is not
     *         a readable JWT
     */
    public AuthMode authMode() {
        if (authMode != AuthMode.AUTH_MODE_UNSPECIFIED) {
            return authMode;
        }
        try {
            return GraphTokens.authMode(bearer.get());
        } catch (RuntimeException e) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
    }

    /**
     * GET {@code path} as JSON with extra request headers (for example
     * {@code Prefer}).
     *
     * @param path Graph path relative to the base URL, or an absolute URL under it
     * @param headers extra request headers; empty for none
     * @return the JSON body; an empty object when the body is blank
     * @throws IOException if Graph returns an error, the body is not JSON, or an
     *         absolute URL is not under the configured Graph base URL
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode get(String path, Map<String, String> headers)
            throws IOException, InterruptedException {
        HttpRequest.Builder builder = request(requireOnGraph(path)).GET();
        headers.forEach(builder::header);
        return json(send(builder.build(), HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * An open binary response: the body stream plus what the server said about it.
     *
     * @param status final HTTP status (200, or 206 for a ranged response)
     * @param contentType {@code Content-Type}; empty when absent
     * @param contentLength {@code Content-Length}; -1 when absent
     * @param body the body stream; the caller closes it
     */
    public record Content(int status, String contentType, long contentLength, InputStream body)
            implements AutoCloseable {

        /**
         * Closes the body stream.
         *
         * @throws IOException if closing fails
         */
        @Override
        public void close() throws IOException {
            body.close();
        }
    }

    /**
     * Opens a binary download. The request to {@code path} carries the bearer token;
     * Graph answers {@code /content} with a 302 to a pre-authenticated URL, which is
     * followed <em>without</em> the {@code Authorization} header (sending the token
     * there would leak it and Graph documents the URL as needing none).
     * {@code range} is sent only to that pre-authenticated URL, as Graph requires;
     * when Graph serves the bytes without a redirect the whole body comes back
     * (status 200) and the caller applies the range itself.
     *
     * @param path Graph path or absolute URL under the base URL
     * @param range an HTTP {@code Range} value such as {@code bytes=0-1023}; empty
     *        for the whole body
     * @return the open response; status 206 when the range was honoured, 200 when
     *         the whole body is returned (the range may be ignored)
     * @throws IOException if Graph returns an error or redirects too often
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public Content openContent(String path, String range)
            throws IOException, InterruptedException {
        HttpRequest first = request(requireOnGraph(path)).GET().build();
        HttpResponse<InputStream> response = send(first, HttpResponse.BodyHandlers.ofInputStream());
        int redirects = 0;
        while (isRedirect(response.statusCode())) {
            String location = response.headers().firstValue("Location").orElse("");
            response.body().close();
            if (location.isEmpty() || ++redirects > MAX_REDIRECTS) {
                throw new GraphApiException(response.statusCode(), "redirect",
                        "download redirect without a usable Location");
            }
            URI target = URI.create(response.uri().toString()).resolve(location);
            HttpRequest.Builder next = HttpRequest.newBuilder(target)
                    .timeout(Duration.ofMinutes(10)).GET();
            if (range != null && !range.isEmpty()) {
                next.header("range", range);
            }
            response = send(next.build(), HttpResponse.BodyHandlers.ofInputStream());
        }
        if (response.statusCode() >= 400) {
            String body;
            try (InputStream in = response.body()) {
                body = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
            }
            throw error(response, body);
        }
        return new Content(response.statusCode(),
                response.headers().firstValue("content-type").orElse(""),
                response.headers().firstValueAsLong("content-length").orElse(-1L),
                response.body());
    }

    private static boolean isRedirect(int status) {
        return status == 301 || status == 302 || status == 303 || status == 307 || status == 308;
    }

    /**
     * Refuses absolute URLs that are not under the configured Graph base URL so a
     * link (from a client or a response) is never fetched with the bearer token
     * on another host.
     */
    private String requireOnGraph(String path) throws IOException {
        if (path.startsWith("http") && !GraphLinks.isUnderBase(baseUrl, path)) {
            throw new IOException("refusing to send the Graph token to a URL outside "
                    + baseUrl);
        }
        return path;
    }

    /**
     * GET {@code path} (relative to the base URL, or an absolute URL) as JSON.
     *
     * @param path Graph path or absolute URL
     * @return the JSON body; an empty object when the body is blank
     * @throws IOException if Graph returns an error or the body is not JSON
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode get(String path) throws IOException, InterruptedException {
        return json(send(request(path).GET().build(), HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * GET {@code path} as raw bytes (file content downloads).
     *
     * @param path Graph path or absolute URL
     * @return the response body bytes
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public byte[] getBytes(String path) throws IOException, InterruptedException {
        try (Content content = openContent(path, "")) {
            return content.body().readAllBytes();
        }
    }

    /**
     * POST JSON to {@code path}.
     *
     * @param path Graph path or absolute URL
     * @param body JSON request body
     * @return the JSON response body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode post(String path, JsonNode body) throws IOException, InterruptedException {
        return json(send(request(path)
                        .header("content-type", "application/json")
                        .POST(HttpRequest.BodyPublishers.ofString(body.toString())).build(),
                HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * POST that answers 202 with a Location header (Graph's async operations).
     *
     * @param path Graph path or absolute URL
     * @param body JSON request body
     * @return the {@code Location} header when present
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public Optional<String> postAsync(String path, JsonNode body)
            throws IOException, InterruptedException {
        HttpResponse<String> response = send(request(path)
                        .header("content-type", "application/json")
                        .POST(HttpRequest.BodyPublishers.ofString(body.toString())).build(),
                HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() >= 400) {
            throw error(response, response.body());
        }
        return response.headers().firstValue("Location");
    }

    /**
     * PATCH JSON at {@code path}.
     *
     * @param path Graph path or absolute URL
     * @param body JSON request body
     * @return the JSON response body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode patch(String path, JsonNode body) throws IOException, InterruptedException {
        return json(send(request(path)
                        .header("content-type", "application/json")
                        .method("PATCH", HttpRequest.BodyPublishers.ofString(body.toString()))
                        .build(),
                HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * PUT JSON at {@code path}.
     *
     * @param path Graph path or absolute URL
     * @param body JSON request body
     * @return the JSON response body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode put(String path, JsonNode body) throws IOException, InterruptedException {
        return json(send(request(path)
                        .header("content-type", "application/json")
                        .PUT(HttpRequest.BodyPublishers.ofString(body.toString())).build(),
                HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * PUT raw bytes at {@code path} with {@code contentType}.
     *
     * @param path Graph path or absolute URL
     * @param content body bytes
     * @param contentType {@code Content-Type} header value
     * @return the JSON response body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putBytes(String path, byte[] content, String contentType)
            throws IOException, InterruptedException {
        return putBytes(path, content, content.length, contentType, Map.of());
    }

    /**
     * PUT the first {@code length} bytes of {@code content} at {@code path}.
     *
     * @param path Graph path or absolute URL
     * @param content body bytes
     * @param length how many bytes of {@code content} to send
     * @param contentType {@code Content-Type} header value
     * @param headers extra request headers; empty for none
     * @return the JSON response body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putBytes(String path, byte[] content, int length, String contentType,
            Map<String, String> headers) throws IOException, InterruptedException {
        HttpRequest.Builder builder = request(path)
                .timeout(Duration.ofMinutes(10))
                .header("content-type", contentType)
                .PUT(HttpRequest.BodyPublishers.ofByteArray(content, 0, length));
        headers.forEach(builder::header);
        return json(send(builder.build(), HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * PUT bytes at an absolute URL with no {@code Authorization} header.
     * Graph upload-session URLs are pre-authenticated; sending a bearer token
     * makes them fail.
     *
     * @param url absolute upload URL from {@code createUploadSession}
     * @param content chunk bytes
     * @param contentType {@code Content-Type}; {@code null} becomes
     *        {@code application/octet-stream}
     * @param contentRange {@code Content-Range} value, for example
     *        {@code bytes 0-3276799/5000000}
     * @return the JSON body, or an empty object when Graph returns no body
     *        (intermediate 202 chunks)
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putRangeUnauthenticated(String url, byte[] content, String contentType,
            String contentRange) throws IOException, InterruptedException {
        return putRangeUnauthenticated(url, content, content.length, contentType, contentRange);
    }

    /**
     * PUT the first {@code length} bytes of {@code content} at an absolute URL with no
     * {@code Authorization} header.
     *
     * @param url absolute upload URL from {@code createUploadSession}
     * @param content chunk bytes
     * @param length how many bytes of {@code content} to send
     * @param contentType {@code Content-Type}; {@code null} becomes
     *        {@code application/octet-stream}
     * @param contentRange {@code Content-Range} value
     * @return the JSON body, or an empty object when Graph returns no body
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putRangeUnauthenticated(String url, byte[] content, int length,
            String contentType, String contentRange) throws IOException, InterruptedException {
        HttpRequest.Builder builder = HttpRequest.newBuilder(URI.create(url))
                .timeout(Duration.ofMinutes(10))
                .header("content-type",
                        contentType == null ? "application/octet-stream" : contentType)
                .PUT(HttpRequest.BodyPublishers.ofByteArray(content, 0, length));
        if (contentRange != null && !contentRange.isBlank()) {
            builder.header("content-range", contentRange);
        }
        HttpResponse<String> response = send(builder.build(), HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() >= 400) {
            throw error(response, response.body());
        }
        String body = response.body();
        if (body == null || body.isBlank()) {
            return JSON.createObjectNode();
        }
        return JSON.readTree(body);
    }

    /**
     * DELETE {@code path}.
     *
     * @param path Graph path or absolute URL
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public void delete(String path) throws IOException, InterruptedException {
        HttpResponse<String> response = send(request(path).DELETE().build(),
                HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() >= 400) {
            throw error(response, response.body());
        }
    }

    /**
     * One Graph request.
     *
     * @param method HTTP method
     * @param path Graph path relative to the base URL
     * @param headers extra request headers
     * @param body JSON body; {@code null} for none
     */
    public record Call(String method, String path, Map<String, String> headers, JsonNode body) {
    }

    /**
     * A successful Graph response.
     *
     * @param status HTTP status
     * @param location {@code Location} header; empty when absent
     * @param body JSON body; an empty object when there was none
     */
    public record Reply(int status, String location, JsonNode body) {
    }

    /**
     * One request of a {@code $batch}.
     *
     * @param call the request
     * @param dependsOn indexes of earlier requests of the batch that must succeed first
     */
    public record BatchCall(Call call, List<Integer> dependsOn) {
    }

    /**
     * The answer to one request of a {@code $batch}; exactly one component is set.
     *
     * @param reply the response when the request succeeded
     * @param error the error when it failed
     */
    public record BatchReply(Reply reply, GraphApiException error) {
    }

    /** Most requests Graph accepts in one {@code $batch}. */
    public static final int MAX_BATCH_SIZE = 20;

    /**
     * Sends one request.
     *
     * @param call the request
     * @return the response
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public Reply execute(Call call) throws IOException, InterruptedException {
        HttpRequest.Builder builder = request(call.path());
        if (call.body() == null) {
            builder.method(call.method(), HttpRequest.BodyPublishers.noBody());
        } else {
            builder.header("content-type", "application/json")
                    .method(call.method(), HttpRequest.BodyPublishers.ofString(call.body().toString()));
        }
        call.headers().forEach(builder::header);
        return reply(send(builder.build(), HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * Sends up to {@link #MAX_BATCH_SIZE} requests as one {@code POST /$batch}.
     *
     * @param calls the requests
     * @return one answer per request, in request order
     * @throws IOException if the batch request itself fails
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public List<BatchReply> batch(List<BatchCall> calls) throws IOException, InterruptedException {
        if (calls.size() > MAX_BATCH_SIZE) {
            throw new IllegalArgumentException("a batch holds at most " + MAX_BATCH_SIZE + " requests");
        }
        ObjectNode envelope = JSON.createObjectNode();
        ArrayNode requests = envelope.putArray("requests");
        for (int i = 0; i < calls.size(); i++) {
            BatchCall batched = calls.get(i);
            Call call = batched.call();
            ObjectNode request = requests.addObject()
                    .put("id", Integer.toString(i))
                    .put("method", call.method())
                    .put("url", call.path());
            ObjectNode headers = JSON.createObjectNode();
            call.headers().forEach(headers::put);
            if (call.body() != null) {
                headers.put("Content-Type", "application/json");
                request.set("body", call.body());
            }
            if (!headers.isEmpty()) {
                request.set("headers", headers);
            }
            if (!batched.dependsOn().isEmpty()) {
                ArrayNode dependsOn = request.putArray("dependsOn");
                batched.dependsOn().forEach(index -> dependsOn.add(Integer.toString(index)));
            }
        }
        JsonNode answer = post("/$batch", envelope);
        BatchReply[] replies = new BatchReply[calls.size()];
        for (JsonNode response : answer.path("responses")) {
            int index;
            try {
                index = Integer.parseInt(response.path("id").asText(""));
            } catch (NumberFormatException e) {
                continue;
            }
            if (index < 0 || index >= replies.length) {
                continue;
            }
            int status = response.path("status").asInt(0);
            JsonNode headers = response.path("headers");
            JsonNode body = response.path("body");
            if (status >= 400) {
                replies[index] = new BatchReply(null, error(status,
                        body.isMissingNode() ? "" : body.toString(), header(headers, "Location"),
                        retryAfter(header(headers, "Retry-After"))));
            } else {
                replies[index] = new BatchReply(new Reply(status, header(headers, "Location"),
                        body.isObject() ? body : JSON.createObjectNode()), null);
            }
        }
        for (int i = 0; i < replies.length; i++) {
            if (replies[i] == null) {
                replies[i] = new BatchReply(null, new GraphApiException(
                        GraphApiException.NO_HTTP_STATUS, "noBatchResponse",
                        "the batch response has no entry for this request"));
            }
        }
        return List.of(replies);
    }

    private static String header(JsonNode headers, String name) {
        for (Map.Entry<String, JsonNode> entry : headers.properties()) {
            if (entry.getKey().equalsIgnoreCase(name)) {
                return entry.getValue().asText("");
            }
        }
        return "";
    }

    /**
     * GET an absolute URL with no {@code Authorization} header (upload-session and
     * operation-monitor URLs).
     *
     * @param url absolute pre-authenticated URL
     * @return the response; redirects are returned, not followed
     * @throws IOException if the server returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public Reply getUnauthenticated(String url) throws IOException, InterruptedException {
        return reply(send(HttpRequest.newBuilder(URI.create(url)).timeout(Duration.ofSeconds(60))
                .GET().build(), HttpResponse.BodyHandlers.ofString()));
    }

    /**
     * DELETE an absolute URL with no {@code Authorization} header (cancels an upload
     * session).
     *
     * @param url absolute pre-authenticated URL
     * @throws IOException if the server returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public void deleteUnauthenticated(String url) throws IOException, InterruptedException {
        reply(send(HttpRequest.newBuilder(URI.create(url)).timeout(Duration.ofSeconds(60))
                .DELETE().build(), HttpResponse.BodyHandlers.ofString()));
    }

    private static Reply reply(HttpResponse<String> response) throws IOException {
        if (response.statusCode() >= 400) {
            throw error(response, response.body());
        }
        JsonNode body;
        try {
            body = parse(response.body());
        } catch (IOException e) {
            body = JSON.createObjectNode();
        }
        return new Reply(response.statusCode(),
                response.headers().firstValue("Location").orElse(""), body);
    }

    /**
     * Follows an async-operation URL until it reports {@code completed} (or fails).
     *
     * @param operationUrl Graph operation monitor URL
     * @param timeout how long to poll before giving up
     * @return the completed operation JSON
     * @throws IOException if Graph reports failure or the wait times out
     * @throws InterruptedException if polling sleep or HTTP is interrupted
     */
    public JsonNode awaitOperation(String operationUrl, Duration timeout)
            throws IOException, InterruptedException {
        long deadline = System.nanoTime() + timeout.toNanos();
        while (true) {
            JsonNode operation = get(operationUrl);
            String status = operation.path("status").asText("");
            if (status.equalsIgnoreCase("completed")) {
                return operation;
            }
            if (status.equalsIgnoreCase("failed")) {
                throw new GraphApiException(GraphApiException.NO_HTTP_STATUS, "operationFailed",
                        operation.path("error").path("message").asText("operation failed"),
                        operation.toString());
            }
            if (System.nanoTime() >= deadline) {
                throw new GraphApiException(GraphApiException.NO_HTTP_STATUS, "operationTimeout",
                        "Async operation still " + status + " after " + timeout);
            }
            Thread.sleep(Duration.ofSeconds(2));
        }
    }

    private HttpRequest.Builder request(String path) {
        String url = path.startsWith("http") ? path : baseUrl + path;
        return HttpRequest.newBuilder(URI.create(url))
                .timeout(Duration.ofSeconds(60))
                .header("authorization", "Bearer " + bearer.get());
    }

    /** Sends with Graph's throttling contract: 429/503 retried after Retry-After. */
    private <T> HttpResponse<T> send(HttpRequest req, HttpResponse.BodyHandler<T> handler)
            throws IOException, InterruptedException {
        HttpResponse<T> response = http.send(req, handler);
        int retries = 0;
        while ((response.statusCode() == 429 || response.statusCode() == 503)
                && retries < MAX_THROTTLE_RETRIES) {
            long waitSeconds = retryAfter(response.headers());
            Thread.sleep(Duration.ofSeconds(Math.min(waitSeconds < 0 ? 2L : waitSeconds, 60)));
            response = http.send(req, handler);
            retries++;
        }
        return response;
    }

    private static JsonNode json(HttpResponse<String> response) throws IOException {
        String body = response.body();
        if (response.statusCode() >= 400) {
            throw error(response, body);
        }
        return parse(body);
    }

    private static JsonNode parse(String body) throws IOException {
        if (body == null || body.isBlank()) {
            return JSON.createObjectNode();
        }
        return JSON.readTree(body);
    }

    /** {@code Retry-After} in seconds; -1 when absent or not a number. */
    private static long retryAfter(HttpHeaders headers) {
        return retryAfter(headers.firstValue("Retry-After").orElse(""));
    }

    private static long retryAfter(String header) {
        String value = header.trim();
        if (value.isEmpty()) {
            return -1L;
        }
        try {
            return Math.max(0L, Long.parseLong(value));
        } catch (NumberFormatException e) {
            return -1L;
        }
    }

    private static GraphApiException error(HttpResponse<?> response, String body) {
        return error(response.statusCode(), body,
                response.headers().firstValue("Location").orElse(""),
                retryAfter(response.headers()));
    }

    static GraphApiException error(int status, String body, String location,
            long retryAfterSeconds) {
        String code = "";
        String message = body == null ? "" : body;
        try {
            JsonNode node = JSON.readTree(body);
            code = node.path("error").path("code").asText("");
            message = node.path("error").path("message").asText(message);
        } catch (Exception ignored) {
            // not JSON; keep the raw body as the message
        }
        return new GraphApiException(status, code, message, body == null ? "" : body, location,
                retryAfterSeconds);
    }

    static ObjectNode object() {
        return JSON.createObjectNode();
    }
}
