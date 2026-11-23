package ai.pipestream.connect;

import ai.pipestream.microsoft.GraphAuth;
import ai.pipestream.microsoft.GraphClient;
import ai.pipestream.microsoft.GraphFiles;
import ai.pipestream.microsoft.InMemoryMicrosoftChangeSink;
import ai.pipestream.microsoft.MicrosoftConnectorConfig;
import ai.pipestream.microsoft.MicrosoftCrawler;
import ai.pipestream.microsoft.SyncCursors;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.Channel;

import java.time.Instant;
import java.time.format.DateTimeParseException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

/**
 * One poll is one MicrosoftService.Sync (gRPC) or one crawler pass
 * (direct). The offset is the pass's resume cursor (an encoded
 * {@code SyncCursor}), and the next poll resumes from it: a delta pass that
 * publishes only what changed since, deletes included. With no offset the
 * pass is a full enumeration.
 */
public final class MicrosoftSourceTask extends ChangeSourceTask<MicrosoftChange> {

    /** Why a pass whose stream had no terminal cursor fails. */
    static final String NO_RESUME_CURSOR = "Sync ended without its resume cursor";

    private static final System.Logger LOG =
            System.getLogger(MicrosoftSourceTask.class.getName());

    /** Kafka Connect task that pulls MicrosoftChange protobuf bytes. */
    public MicrosoftSourceTask() {
        super("microsoft");
    }

    private boolean includeContent;
    private List<String> driveIds = List.of();
    private String folderPath = "/";
    private MicrosoftConnectorConfig direct;

    @Override
    protected void configure(Map<String, String> props, boolean directMode) {
        includeContent = Boolean.parseBoolean(
                props.getOrDefault(MicrosoftSourceConnector.INCLUDE_CONTENT, "false"));
        driveIds = splitCsv(props.getOrDefault(MicrosoftSourceConnector.DRIVE_IDS, ""));
        folderPath = props.getOrDefault(MicrosoftSourceConnector.FOLDER_PATH, "/");
        if (directMode) {
            MicrosoftConnectorConfig.Builder builder = MicrosoftConnectorConfig.builder()
                    .tenantId(required(props, MicrosoftSourceConnector.TENANT_ID))
                    .clientId(required(props, MicrosoftSourceConnector.CLIENT_ID))
                    .clientSecret(required(props, MicrosoftSourceConnector.CLIENT_SECRET))
                    .siteId(props.getOrDefault(MicrosoftSourceConnector.SITE_ID, ""))
                    .driveIds(driveIds)
                    .folderPath(folderPath);
            String graph = props.getOrDefault(MicrosoftSourceConnector.GRAPH_BASE_URL, "");
            if (!graph.isBlank()) {
                builder.graphBaseUrl(graph);
            }
            String authority = props.getOrDefault(MicrosoftSourceConnector.AUTHORITY, "");
            if (!authority.isBlank()) {
                builder.authority(authority);
            }
            direct = builder.build();
        }
    }

    @Override
    protected String pollGrpc(Channel channel, String lastCursor, List<MicrosoftChange> out) {
        MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                MicrosoftServiceGrpc.newBlockingStub(channel);
        SyncRequest.Builder request = SyncRequest.newBuilder()
                .setIncludeContent(includeContent)
                .setSinceCursor(sinceCursor(lastCursor));
        request.addAllDriveIds(driveIds).setFolderPath(folderPath);
        Iterator<SyncResponse> events = stub.sync(request.build());
        String cursor = "";
        while (events.hasNext()) {
            SyncResponse event = events.next();
            if (event.hasChange()) {
                out.add(event.getChange());
            } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                cursor = event.getResumeCursor();
            }
        }
        if (cursor.isEmpty()) {
            // Every Sync stream ends with its resume cursor; one without it
            // was cut short, and committing it would lose what was not sent.
            throw new IllegalStateException(NO_RESUME_CURSOR);
        }
        return cursor;
    }

    @Override
    protected String pollDirect(String lastCursor, List<MicrosoftChange> out) throws Exception {
        SyncCursor since = SyncCursors.decode(sinceCursor(lastCursor), direct.graphBaseUrl());
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        GraphAuth auth = new GraphAuth(direct.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(auth.clientCredentials());
        GraphFiles files = new GraphFiles(new GraphClient(direct.graphBaseUrl(), () -> {
            GraphAuth.Token current = token.get();
            if (current.expired()) {
                try {
                    current = auth.clientCredentials();
                    token.set(current);
                } catch (Exception e) {
                    throw new IllegalStateException("token refresh failed", e);
                }
            }
            return current.accessToken();
        }));
        MicrosoftCrawler.Result result = new MicrosoftCrawler(direct, files, sink,
                MicrosoftCrawler.Options.defaults(direct, false, 0).withSince(since)).crawl();
        out.addAll(sink.changes());
        return SyncCursors.encode(result.cursor());
    }

    /**
     * The stored offset as a {@code since_cursor}. Versions before the
     * resume cursor was sent back stored a wall-clock ISO-8601 instant as
     * the offset; that was never a cursor Sync issued, so it reads as "no
     * cursor" and the next pass is the full enumeration that starts delta
     * tracking. Anything else is passed on as stored, and Sync refuses what
     * it did not issue.
     *
     * @param lastCursor the stored offset, or {@code null}
     * @return the cursor to resume from; empty asks for a full enumeration
     */
    static String sinceCursor(String lastCursor) {
        if (lastCursor == null || lastCursor.isBlank()) {
            return "";
        }
        try {
            Instant.parse(lastCursor);
        } catch (DateTimeParseException notAnInstant) {
            return lastCursor;
        }
        LOG.log(System.Logger.Level.WARNING,
                "stored offset {0} is a wall-clock time from an older version, not a resume"
                        + " cursor: this pass is a full enumeration and reports no deletes",
                lastCursor);
        return "";
    }

    @Override
    protected String recordKey(MicrosoftChange change) {
        return change.getEntity().getEntityId();
    }
}
