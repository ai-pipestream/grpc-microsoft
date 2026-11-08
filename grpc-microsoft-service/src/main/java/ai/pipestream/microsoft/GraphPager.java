package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DeltaResync;
import ai.pipestream.microsoft.v1.DeltaResyncReason;
import com.fasterxml.jackson.databind.JsonNode;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Hides Graph paging: walks {@code @odata.nextLink} for collections and runs
 * delta enumerations to their {@code @odata.deltaLink}, turning HTTP 410 Gone into
 * a typed {@link DeltaResync} followed by the fresh enumeration Graph's
 * {@code Location} header points at. Every link Graph returns is checked to be
 * under the configured Graph base URL before the token is sent to it.
 *
 * <p>Each request may name several candidate first URLs: the preferred one
 * (with the full {@code $select} / {@code $expand}) and progressively plainer
 * fallbacks. A {@code 400 Bad Request} on a candidate's first page moves on to
 * the next one, so a query option a Graph endpoint rejects degrades to Graph's
 * default property set instead of failing the rpc.</p>
 */
public final class GraphPager {

    /** Callback for one collection element; return {@code false} to stop paging. */
    @FunctionalInterface
    public interface Visitor {
        /**
         * Receives one element of {@code value}.
         *
         * @param item the element
         * @return {@code true} to keep going
         * @throws IOException if handling needs Graph and Graph fails
         * @throws InterruptedException if handling is interrupted
         */
        boolean visit(JsonNode item) throws IOException, InterruptedException;
    }

    /** Callbacks for a delta enumeration. */
    public interface DeltaVisitor {
        /**
         * Graph rejected the delta link (410 Gone); a fresh full enumeration follows.
         *
         * @param resync why, as Graph reported it
         * @throws IOException if handling fails
         * @throws InterruptedException if handling is interrupted
         */
        void resync(DeltaResync resync) throws IOException, InterruptedException;

        /**
         * One changed (or deleted) element.
         *
         * @param item the element
         * @throws IOException if handling needs Graph and Graph fails
         * @throws InterruptedException if handling is interrupted
         */
        void item(JsonNode item) throws IOException, InterruptedException;
    }

    /**
     * Result of a delta enumeration.
     *
     * @param deltaLink the terminal {@code @odata.deltaLink}
     * @param resynced whether Graph rejected the starting link (410) and the
     *        enumeration was a fresh full one
     */
    public record DeltaOutcome(String deltaLink, boolean resynced) {
    }

    /**
     * Graph handed back a {@code @odata.nextLink} it had already returned in
     * the same enumeration: following it would loop forever.
     */
    public static final class RepeatedLinkException extends IOException {
        /** Serialization identifier. */
        @java.io.Serial
        private static final long serialVersionUID = 1L;

        RepeatedLinkException(String link) {
            super("Graph repeated the next link " + link + "; refusing to loop");
        }
    }

    /**
     * Admits the next link of one enumeration: {@link RepeatedLinkException}
     * when it was already followed, {@code CANCELLED} when the gRPC client
     * that the enumeration serves has gone.
     *
     * @param seen links already followed in this enumeration
     * @param link the next link
     * @return {@code link}
     * @throws RepeatedLinkException when {@code link} repeats
     */
    static String admit(java.util.Set<String> seen, String link) throws RepeatedLinkException {
        if (!seen.add(link)) {
            throw new RepeatedLinkException(link);
        }
        if (io.grpc.Context.current().isCancelled()) {
            throw io.grpc.Status.CANCELLED.withDescription("the client cancelled the call")
                    .asRuntimeException();
        }
        return link;
    }

    private final GraphClient graph;

    /**
     * Binds a pager to a client.
     *
     * @param graph the authorized client
     */
    public GraphPager(GraphClient graph) {
        this.graph = Objects.requireNonNull(graph, "graph");
    }

    /**
     * GETs one entity, trying each candidate URL in turn on {@code 400}.
     *
     * @param candidates first-choice URL, then fallbacks
     * @param headers request headers
     * @return the JSON body
     * @throws IOException if Graph fails
     * @throws InterruptedException if interrupted
     */
    public JsonNode get(List<String> candidates, Map<String, String> headers)
            throws IOException, InterruptedException {
        for (int i = 0; ; i++) {
            try {
                return graph.get(candidates.get(i), headers);
            } catch (GraphClient.GraphApiException e) {
                if (e.status() != 400 || i == candidates.size() - 1) {
                    throw e;
                }
            }
        }
    }

    /**
     * Streams every element of a collection across pages.
     *
     * @param candidates first-choice URL, then fallbacks
     * @param headers request headers sent on every page
     * @param visitor receives each element
     * @throws IOException if Graph fails or a next link leaves the Graph host
     * @throws InterruptedException if interrupted
     */
    public void forEach(List<String> candidates, Map<String, String> headers, Visitor visitor)
            throws IOException, InterruptedException {
        JsonNode page = get(candidates, headers);
        java.util.Set<String> seen = new java.util.HashSet<>();
        while (true) {
            for (JsonNode item : page.path("value")) {
                if (!visitor.visit(item)) {
                    return;
                }
            }
            String next = page.path("@odata.nextLink").asText("");
            if (next.isBlank()) {
                return;
            }
            page = graph.get(admit(seen, next), headers);
        }
    }

    /**
     * Runs one delta enumeration to its delta link.
     *
     * @param fresh candidate URLs for a fresh enumeration (no token)
     * @param startLink an already-validated delta link to resume from, or empty
     *        for a fresh enumeration
     * @param headers request headers ({@code Prefer}) sent on every page
     * @param visitor receives the resync notice and every element
     * @return the terminal delta link and whether a resync happened
     * @throws IOException if Graph fails, a page has neither link, or Graph
     *         rejects the fresh enumeration too
     * @throws InterruptedException if interrupted
     */
    public DeltaOutcome delta(List<String> fresh, String startLink, Map<String, String> headers,
            DeltaVisitor visitor) throws IOException, InterruptedException {
        boolean resynced = false;
        String next = startLink == null || startLink.isEmpty() ? null : startLink;
        java.util.Set<String> seen = new java.util.HashSet<>();
        while (true) {
            JsonNode page;
            try {
                if (next == null) {
                    page = get(fresh, headers);
                } else {
                    page = graph.get(next, headers);
                }
            } catch (GraphClient.GraphApiException e) {
                if (e.status() != 410 || resynced) {
                    throw e;
                }
                resynced = true;
                visitor.resync(DeltaResync.newBuilder()
                        .setReason(reason(e))
                        .setMessage(e.getMessage())
                        .build());
                next = e.location().isEmpty() ? null : e.location();
                continue;
            }
            for (JsonNode item : page.path("value")) {
                visitor.item(item);
            }
            String nextLink = page.path("@odata.nextLink").asText("");
            if (!nextLink.isBlank()) {
                next = admit(seen, nextLink);
                continue;
            }
            String deltaLink = page.path("@odata.deltaLink").asText("");
            if (deltaLink.isBlank()) {
                throw new IOException("Graph delta page carried neither @odata.nextLink nor "
                        + "@odata.deltaLink");
            }
            return new DeltaOutcome(deltaLink, resynced);
        }
    }

    static DeltaResyncReason reason(GraphClient.GraphApiException e) {
        String code = e.code() + " " + e.body();
        if (code.contains("resyncChangesUploadDifferences")) {
            return DeltaResyncReason.DELTA_RESYNC_REASON_UPLOAD_DIFFERENCES;
        }
        return DeltaResyncReason.DELTA_RESYNC_REASON_APPLY_DIFFERENCES;
    }
}
