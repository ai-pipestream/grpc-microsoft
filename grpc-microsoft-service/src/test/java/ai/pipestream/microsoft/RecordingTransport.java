package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Sends requests to the real server and keeps every exchange, scrubbed, until
 * the scenario that made them is saved or dropped.
 */
final class RecordingTransport implements GraphTransport {

    private final GraphTransport real;
    private final RecordingScrubber scrubber;
    private final Map<String, byte[]> blobs;
    private final String graphBaseUrl;
    private final List<ObjectNode> exchanges = new ArrayList<>();

    /**
     * Creates a transport that records what {@code real} exchanges.
     *
     * @param real sends the requests
     * @param scrubber replaces tenant data
     * @param blobs content the scenarios produced, by sha256
     * @param graphBaseUrl the Graph base URL
     */
    RecordingTransport(GraphTransport real, RecordingScrubber scrubber, Map<String, byte[]> blobs,
            String graphBaseUrl) {
        this.real = real;
        this.scrubber = scrubber;
        this.blobs = blobs;
        this.graphBaseUrl = graphBaseUrl;
    }

    /** A transport that records what a new {@link HttpClient} exchanges. */
    static RecordingTransport overHttp(RecordingScrubber scrubber, Map<String, byte[]> blobs,
            String graphBaseUrl) {
        return new RecordingTransport(
                HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(20)).build()::send,
                scrubber, blobs, graphBaseUrl);
    }

    @Override
    public <T> HttpResponse<T> send(HttpRequest request, HttpResponse.BodyHandler<T> handler)
            throws IOException, InterruptedException {
        byte[] sent = Recordings.body(request);
        HttpResponse<byte[]> response = real.send(request, HttpResponse.BodyHandlers.ofByteArray());
        byte[] received = response.body() == null ? new byte[0] : response.body();
        synchronized (this) {
            ObjectNode exchange = Recordings.JSON.createObjectNode();
            exchange.set("request", Recordings.request(request, sent, scrubber));
            exchange.set("response", Recordings.response(response.statusCode(), response.headers(),
                    received, scrubber, blobs));
            exchanges.add(exchange);
        }
        return Recordings.respond(request, response.statusCode(), response.headers(), received,
                handler);
    }

    /** Forgets the exchanges recorded so far. */
    synchronized void drop() {
        exchanges.clear();
    }

    /**
     * Writes the exchanges recorded so far to {@code scenarioDir}, one file each,
     * and forgets them.
     *
     * @return how many files were written
     * @throws IllegalStateException when an exchange still holds tenant data
     */
    synchronized int save(Path scenarioDir) throws IOException {
        Files.createDirectories(scenarioDir);
        int number = 0;
        for (ObjectNode exchange : exchanges) {
            number++;
            String text = Recordings.JSON.writerWithDefaultPrettyPrinter()
                    .writeValueAsString(exchange) + "\n";
            List<String> violations = RecordingScrubber.violations(text);
            if (!violations.isEmpty()) {
                throw new IllegalStateException("exchange " + number + " of "
                        + scenarioDir.getFileName() + " still holds tenant data: " + violations);
            }
            String url = exchange.path("request").path("url").asText();
            String name = String.format(Locale.ROOT, "%03d-%s-%s.json", number,
                    exchange.path("request").path("method").asText(),
                    Recordings.shortRoute(url, graphBaseUrl));
            Files.writeString(scenarioDir.resolve(name), text, StandardCharsets.UTF_8);
        }
        exchanges.clear();
        return number;
    }
}
