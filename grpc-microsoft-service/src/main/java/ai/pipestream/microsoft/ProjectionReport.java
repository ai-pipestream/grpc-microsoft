package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DegradedProjection;
import ai.pipestream.microsoft.v1.DegradedProjections;
import io.grpc.Context;
import io.grpc.Contexts;
import io.grpc.ForwardingServerCall;
import io.grpc.Metadata;
import io.grpc.ServerCall;
import io.grpc.ServerCallHandler;
import io.grpc.ServerInterceptor;
import io.grpc.Status;
import io.grpc.protobuf.ProtoUtils;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * The reads of one rpc that Graph answered from its default property set
 * after rejecting the query the server asked with ({@link GraphPager}). Such
 * a read succeeds, but its entities lack every property only the rejected
 * query asks for, so it is never silent: each distinct one is logged at
 * WARNING, and the call ends with a {@link DegradedProjections} message in
 * its trailing metadata under {@link #TRAILER}.
 *
 * <p>The report of the current call travels in the gRPC {@link Context}
 * ({@link #interceptor()} puts it there). A read made outside an rpc, such
 * as a crawl run in-process, has no report: it is logged once per rejected
 * query for the life of the process.</p>
 */
public final class ProjectionReport {

    /** Trailing metadata key of a call's degraded reads. */
    public static final Metadata.Key<DegradedProjections> TRAILER =
            ProtoUtils.keyForProto(DegradedProjections.getDefaultInstance());

    /**
     * Distinct degraded reads listed in the trailer. Trailers share the
     * client's metadata size limit (8 KiB by default), and a trailer over
     * it fails a call that succeeded.
     */
    static final int MAX_LISTED = 4;
    /** Longest query or message kept in the trailer. */
    static final int MAX_TEXT = 200;

    private static final System.Logger LOG = System.getLogger(ProjectionReport.class.getName());
    /** The report of the call the current context belongs to; absent outside an rpc. */
    static final Context.Key<ProjectionReport> CURRENT =
            Context.key("ai.pipestream.microsoft.projection-report");
    /** Rejected queries already logged by reads made outside any rpc. */
    private static final Set<String> LOGGED_OUTSIDE_A_CALL = ConcurrentHashMap.newKeySet();

    /** Degraded reads by rejected query and used query, in the order first seen. */
    private final Map<String, DegradedProjection.Builder> degraded = new LinkedHashMap<>();

    ProjectionReport() {
    }

    /**
     * Records that a read was answered by a fallback candidate.
     *
     * @param rejectedUrl the candidate Graph rejected (the preferred one)
     * @param usedUrl the candidate the answer came from
     * @param rejection Graph's rejection of {@code rejectedUrl}
     */
    static void degraded(String rejectedUrl, String usedUrl,
            GraphClient.GraphApiException rejection) {
        ProjectionReport report = CURRENT.get();
        boolean first = report != null
                ? report.add(rejectedUrl, usedUrl, rejection)
                : LOGGED_OUTSIDE_A_CALL.add(query(rejectedUrl));
        if (first) {
            LOG.log(System.Logger.Level.WARNING, "Graph rejected the query of {0} ({1}); answered "
                    + "from {2}, so the properties only the rejected query asks for are missing: {3}",
                    route(rejectedUrl), rejection.getMessage(),
                    query(usedUrl).isEmpty() ? "its default properties" : "?" + query(usedUrl),
                    query(rejectedUrl));
        }
    }

    /** Adds one degraded read; whether it is the first of its kind in this call. */
    private synchronized boolean add(String rejectedUrl, String usedUrl,
            GraphClient.GraphApiException rejection) {
        String key = query(rejectedUrl) + "\n" + query(usedUrl);
        DegradedProjection.Builder known = degraded.get(key);
        if (known != null) {
            known.setReads(known.getReads() + 1);
            return false;
        }
        degraded.put(key, DegradedProjection.newBuilder()
                .setRoute(route(rejectedUrl))
                .setRejectedQuery(cut(query(rejectedUrl)))
                .setUsedQuery(cut(query(usedUrl)))
                .setGraphCode(cut(rejection.code()))
                .setGraphMessage(cut(rejection.getMessage()))
                .setReads(1));
        return true;
    }

    /**
     * The degraded reads recorded so far.
     *
     * @return the message for the trailer; its default instance when every
     *         read was answered in full
     */
    synchronized DegradedProjections build() {
        DegradedProjections.Builder out = DegradedProjections.newBuilder();
        List<DegradedProjection.Builder> all = new ArrayList<>(degraded.values());
        for (int i = 0; i < Math.min(all.size(), MAX_LISTED); i++) {
            out.addProjections(all.get(i));
        }
        return out.setOmitted(Math.max(0, all.size() - MAX_LISTED)).build();
    }

    /**
     * The interceptor that gives every call a report and writes it to the
     * call's trailing metadata.
     *
     * @return the interceptor
     */
    static ServerInterceptor interceptor() {
        return new ServerInterceptor() {
            @Override
            public <Q, R> ServerCall.Listener<Q> interceptCall(ServerCall<Q, R> call,
                    Metadata headers, ServerCallHandler<Q, R> next) {
                ProjectionReport report = new ProjectionReport();
                ServerCall<Q, R> reporting =
                        new ForwardingServerCall.SimpleForwardingServerCall<>(call) {
                            @Override
                            public void close(Status status, Metadata trailers) {
                                DegradedProjections degraded = report.build();
                                if (degraded.getProjectionsCount() > 0) {
                                    trailers.put(TRAILER, degraded);
                                }
                                super.close(status, trailers);
                            }
                        };
                return Contexts.interceptCall(Context.current().withValue(CURRENT, report),
                        reporting, headers, next);
            }
        };
    }

    /** The route of a candidate URL: its path, without the query and the Graph base. */
    static String route(String url) {
        int query = url.indexOf('?');
        String bare = query < 0 ? url : url.substring(0, query);
        int scheme = bare.indexOf("://");
        if (scheme < 0) {
            return bare;
        }
        int path = bare.indexOf('/', scheme + 3);
        return path < 0 ? "/" : bare.substring(path);
    }

    /** The query of a candidate URL, without the question mark; empty when it has none. */
    static String query(String url) {
        int query = url.indexOf('?');
        return query < 0 ? "" : url.substring(query + 1);
    }

    private static String cut(String text) {
        return text.length() <= MAX_TEXT ? text : text.substring(0, MAX_TEXT - 3) + "...";
    }
}
