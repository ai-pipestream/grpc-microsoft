package ai.pipestream.connect;

import ai.pipestream.microsoft.GraphAuth;
import ai.pipestream.microsoft.GraphClient;
import ai.pipestream.microsoft.GraphFiles;
import ai.pipestream.microsoft.InMemoryMicrosoftChangeSink;
import ai.pipestream.microsoft.MicrosoftConnectorConfig;
import ai.pipestream.microsoft.MicrosoftCrawler;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.Channel;

import java.time.Instant;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

/**
 * One poll is one MicrosoftService.Sync (gRPC) or one crawler pass
 * (direct). Offset is the resume cursor.
 */
public final class MicrosoftSourceTask extends ChangeSourceTask<MicrosoftChange> {

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
            direct = builder.build();
        }
    }

    @Override
    protected String pollGrpc(Channel channel, String lastCursor, List<MicrosoftChange> out) {
        MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                MicrosoftServiceGrpc.newBlockingStub(channel);
        SyncRequest.Builder request = SyncRequest.newBuilder()
                .setIncludeContent(includeContent);
        request.addAllDriveIds(driveIds).setFolderPath(folderPath);
        Iterator<SyncResponse> events = stub.sync(request.build());
        String cursor = fallbackCursor(lastCursor);
        while (events.hasNext()) {
            SyncResponse event = events.next();
            if (event.hasChange()) {
                out.add(event.getChange());
            } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                cursor = event.getResumeCursor();
            }
        }
        return cursor;
    }

    @Override
    protected String pollDirect(String lastCursor, List<MicrosoftChange> out) throws Exception {
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
        new MicrosoftCrawler(direct, files, sink).crawl();
        out.addAll(sink.changes());
        return Instant.now().toString();
    }

    @Override
    protected String recordKey(MicrosoftChange change) {
        return change.getEntity().getEntityId();
    }
}
