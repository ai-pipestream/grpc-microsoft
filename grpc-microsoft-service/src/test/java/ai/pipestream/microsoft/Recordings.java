package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpHeaders;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.ByteBuffer;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HexFormat;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.Flow;
import java.util.stream.Stream;

/** The on-disk form of recorded Graph exchanges, shared by recording and replay. */
final class Recordings {

    /** Where the recordings live, relative to the module directory. */
    static final String DEFAULT_DIR = "src/test/resources/recordings/microsoft";
    /** System property that overrides {@link #DEFAULT_DIR}. */
    static final String DIR_PROPERTY = "recordings.dir";
    /** File in the recordings directory that describes the recorded run. */
    static final String RUN_FILE = "run.json";

    static final ObjectMapper JSON = new ObjectMapper();

    /** Request headers that are recorded and must match on replay. */
    private static final List<String> REQUEST_HEADERS = List.of("if-match", "content-range",
            "content-type", "prefer", "range");
    /** Response headers that are recorded and served on replay. */
    private static final List<String> RESPONSE_HEADERS = List.of("etag", "location",
            "retry-after", "content-range", "content-type");

    private Recordings() {
    }

    /** The recordings directory of this build. */
    static Path dir() {
        return Path.of(System.getProperty(DIR_PROPERTY, DEFAULT_DIR));
    }

    /** The exchange files of one scenario, in order. */
    static List<Path> exchanges(Path scenarioDir) throws IOException {
        try (Stream<Path> files = Files.list(scenarioDir)) {
            return files.filter(file -> file.getFileName().toString().endsWith(".json")).sorted()
                    .toList();
        }
    }

    static String sha256(byte[] bytes) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes));
        } catch (NoSuchAlgorithmException e) {
            throw new IllegalStateException(e);
        }
    }

    /** The bytes a request would send. */
    static byte[] body(HttpRequest request) {
        Optional<HttpRequest.BodyPublisher> publisher = request.bodyPublisher();
        if (publisher.isEmpty() || publisher.get().contentLength() == 0) {
            return new byte[0];
        }
        java.io.ByteArrayOutputStream out = new java.io.ByteArrayOutputStream();
        publisher.get().subscribe(new Flow.Subscriber<>() {
            @Override
            public void onSubscribe(Flow.Subscription subscription) {
                subscription.request(Long.MAX_VALUE);
            }

            @Override
            public void onNext(ByteBuffer item) {
                byte[] part = new byte[item.remaining()];
                item.get(part);
                out.writeBytes(part);
            }

            @Override
            public void onError(Throwable throwable) {
                throw new IllegalStateException(throwable);
            }

            @Override
            public void onComplete() {
            }
        });
        return out.toByteArray();
    }

    /**
     * The recorded form of a request.
     *
     * @param scrubber replaces tenant data; {@code null} to keep the request as sent
     */
    static ObjectNode request(HttpRequest request, byte[] body, RecordingScrubber scrubber) {
        ObjectNode node = JSON.createObjectNode();
        node.put("method", request.method());
        node.put("url", scrub(scrubber, request.uri().toString()));
        node.put("authorized", request.headers().firstValue("authorization").isPresent());
        ObjectNode headers = node.putObject("headers");
        for (String name : REQUEST_HEADERS) {
            request.headers().firstValue(name)
                    .ifPresent(value -> headers.put(name, scrub(scrubber, value)));
        }
        if (body.length > 0) {
            ObjectNode recorded = node.putObject("body");
            JsonNode json = request.headers().firstValue("content-type").orElse("")
                    .startsWith("application/json") ? parse(body) : null;
            if (json != null) {
                recorded.set("json", scrubber == null ? json : scrubber.json(json));
            } else {
                recorded.put("sha256", sha256(body)).put("length", body.length);
            }
        }
        return node;
    }

    /** The recorded form of a response. */
    static ObjectNode response(int status, HttpHeaders headers, byte[] body,
            RecordingScrubber scrubber, Map<String, byte[]> blobs) {
        ObjectNode node = JSON.createObjectNode();
        node.put("status", status);
        ObjectNode recordedHeaders = node.putObject("headers");
        for (String name : RESPONSE_HEADERS) {
            headers.firstValue(name)
                    .ifPresent(value -> recordedHeaders.put(name, scrubber.text(value)));
        }
        if (body.length > 0) {
            ObjectNode recorded = node.putObject("body");
            String digest = sha256(body);
            JsonNode json = headers.firstValue("content-type").orElse("").contains("json")
                    ? parse(body) : null;
            if (json != null) {
                recorded.set("json", scrubber.json(json));
            } else if (blobs.containsKey(digest)) {
                recorded.put("blob", digest).put("length", body.length);
            } else {
                String text = utf8(body);
                if (text != null) {
                    recorded.put("text", scrubber.text(text));
                } else {
                    recorded.put("base64", Base64.getEncoder().encodeToString(body));
                }
            }
        }
        return node;
    }

    /** The bytes of a recorded response body. */
    static byte[] responseBody(JsonNode response, Map<String, byte[]> blobs) throws IOException {
        JsonNode body = response.path("body");
        if (body.has("json")) {
            return JSON.writeValueAsBytes(body.get("json"));
        }
        if (body.has("text")) {
            return body.get("text").asText().getBytes(StandardCharsets.UTF_8);
        }
        if (body.has("base64")) {
            return Base64.getDecoder().decode(body.get("base64").asText());
        }
        if (body.has("blob")) {
            byte[] blob = blobs.get(body.get("blob").asText());
            if (blob == null) {
                throw new IOException("the scenario produced no content with sha256 "
                        + body.get("blob").asText());
            }
            return blob;
        }
        return new byte[0];
    }

    /** A response with the given status, headers and body, read through {@code handler}. */
    static <T> HttpResponse<T> respond(HttpRequest request, int status, HttpHeaders headers,
            byte[] body, HttpResponse.BodyHandler<T> handler) {
        HttpResponse.BodySubscriber<T> subscriber = handler.apply(new HttpResponse.ResponseInfo() {
            @Override
            public int statusCode() {
                return status;
            }

            @Override
            public HttpHeaders headers() {
                return headers;
            }

            @Override
            public HttpClient.Version version() {
                return HttpClient.Version.HTTP_1_1;
            }
        });
        subscriber.onSubscribe(new Flow.Subscription() {
            @Override
            public void request(long n) {
            }

            @Override
            public void cancel() {
            }
        });
        if (body.length > 0) {
            subscriber.onNext(List.of(ByteBuffer.wrap(body)));
        }
        subscriber.onComplete();
        T value = subscriber.getBody().toCompletableFuture().join();
        return new HttpResponse<>() {
            @Override
            public int statusCode() {
                return status;
            }

            @Override
            public HttpRequest request() {
                return request;
            }

            @Override
            public Optional<HttpResponse<T>> previousResponse() {
                return Optional.empty();
            }

            @Override
            public HttpHeaders headers() {
                return headers;
            }

            @Override
            public T body() {
                return value;
            }

            @Override
            public Optional<javax.net.ssl.SSLSession> sslSession() {
                return Optional.empty();
            }

            @Override
            public java.net.URI uri() {
                return request.uri();
            }

            @Override
            public HttpClient.Version version() {
                return HttpClient.Version.HTTP_1_1;
            }
        };
    }

    /** Headers from a recorded {@code headers} object. */
    static HttpHeaders headers(JsonNode recorded) {
        Map<String, List<String>> map = new LinkedHashMap<>();
        for (Map.Entry<String, JsonNode> entry : recorded.properties()) {
            map.put(entry.getKey(), List.of(entry.getValue().asText()));
        }
        return HttpHeaders.of(map, (name, value) -> true);
    }

    /** A short file-name form of a recorded URL, without ids. */
    static String shortRoute(String url, String graphBaseUrl) {
        if (!url.startsWith(graphBaseUrl)) {
            return "external";
        }
        String path = url.substring(graphBaseUrl.length());
        int query = path.indexOf('?');
        if (query >= 0) {
            path = path.substring(0, query);
        }
        List<String> kept = new ArrayList<>();
        for (String segment : path.split("/")) {
            String lower = segment.toLowerCase(Locale.ROOT);
            boolean id = lower.contains("drive-") || lower.contains("item-")
                    || lower.contains("tenant-") || lower.contains(RecordingScrubber.GUID_PREFIX)
                    || (!lower.isEmpty() && lower.chars().allMatch(Character::isDigit));
            if (!segment.isEmpty() && !id) {
                kept.add(lower);
            }
        }
        StringBuilder name = new StringBuilder();
        for (char c : String.join("-", kept).toCharArray()) {
            boolean plain = (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '.';
            if (plain) {
                name.append(c);
            } else if (name.length() > 0 && name.charAt(name.length() - 1) != '-') {
                name.append('-');
            }
        }
        String route = name.length() > 48 ? name.substring(0, 48) : name.toString();
        return route.isEmpty() ? "root" : route;
    }

    private static String scrub(RecordingScrubber scrubber, String value) {
        return scrubber == null ? value : scrubber.text(value);
    }

    private static JsonNode parse(byte[] body) {
        try {
            return JSON.readTree(body);
        } catch (IOException e) {
            return null;
        }
    }

    private static String utf8(byte[] body) {
        try {
            return StandardCharsets.UTF_8.newDecoder().onMalformedInput(CodingErrorAction.REPORT)
                    .onUnmappableCharacter(CodingErrorAction.REPORT).decode(ByteBuffer.wrap(body))
                    .toString();
        } catch (CharacterCodingException e) {
            return null;
        }
    }
}
