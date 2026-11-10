package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Serves recorded exchanges in order and refuses any request that is not the
 * next recorded one.
 */
final class ReplayTransport implements GraphTransport {

    private final Map<String, byte[]> blobs;
    private final List<String> problems = new ArrayList<>();
    private List<Path> files = List.of();
    private int next;

    /**
     * Creates a transport with nothing to serve yet.
     *
     * @param blobs content the scenarios produced, by sha256
     */
    ReplayTransport(Map<String, byte[]> blobs) {
        this.blobs = blobs;
    }

    /** Starts serving the exchanges of {@code scenarioDir}. */
    synchronized void begin(Path scenarioDir) throws IOException {
        files = Recordings.exchanges(scenarioDir);
        next = 0;
        problems.clear();
    }

    /** What went wrong since {@link #begin}: mismatches, and exchanges never asked for. */
    synchronized List<String> finish() {
        List<String> all = new ArrayList<>(problems);
        if (next < files.size()) {
            all.add((files.size() - next) + " recorded exchange(s) were never requested, from "
                    + files.get(next).getFileName());
        }
        files = List.of();
        next = 0;
        problems.clear();
        return all;
    }

    @Override
    public synchronized <T> HttpResponse<T> send(HttpRequest request,
            HttpResponse.BodyHandler<T> handler) throws IOException {
        ObjectNode actual = Recordings.request(request, Recordings.body(request), null);
        if (next >= files.size()) {
            throw mismatch("no recorded exchange is left for " + actual);
        }
        Path file = files.get(next);
        JsonNode exchange = Recordings.JSON.readTree(file.toFile());
        JsonNode expected = exchange.path("request");
        if (!expected.equals(actual)) {
            throw mismatch(file.getFileName() + " expects " + expected + " but the request was "
                    + actual);
        }
        next++;
        JsonNode response = exchange.path("response");
        ObjectNode headers = response.path("headers").deepCopy();
        if (headers.has("retry-after")) {
            headers.put("retry-after", "0");
        }
        return Recordings.respond(request, response.path("status").asInt(),
                Recordings.headers(headers), Recordings.responseBody(response, blobs), handler);
    }

    private IOException mismatch(String message) {
        problems.add(message);
        return new IOException("replay mismatch: " + message);
    }
}
