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
import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;
import org.apache.kafka.connect.data.Schema;
import org.apache.kafka.connect.source.SourceRecord;
import org.apache.kafka.connect.source.SourceTask;

import java.time.Instant;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;

/**
 * One poll is one MicrosoftService.Sync (gRPC) or one crawler pass
 * (direct). Offset is the resume cursor.
 */
public final class MicrosoftSourceTask extends SourceTask {

    static final Map<String, String> PARTITION = Map.of("source", "microsoft");

    private String topic;
    private String grpcTarget;
    private boolean plaintext;
    private boolean includeContent;
    private List<String> driveIds = List.of();
    private String folderPath = "/";
    private MicrosoftConnectorConfig direct;
    private String lastCursor;

    @Override
    public String version() {
        return ConnectVersion.VALUE;
    }

    @Override
    public void start(Map<String, String> props) {
        topic = required(props, MicrosoftSourceConnector.TOPIC);
        grpcTarget = props.getOrDefault(MicrosoftSourceConnector.GRPC_TARGET, "").trim();
        plaintext = Boolean.parseBoolean(
                props.getOrDefault(MicrosoftSourceConnector.GRPC_PLAINTEXT, "true"));
        includeContent = Boolean.parseBoolean(
                props.getOrDefault(MicrosoftSourceConnector.INCLUDE_CONTENT, "false"));
        driveIds = splitCsv(props.getOrDefault(MicrosoftSourceConnector.DRIVE_IDS, ""));
        folderPath = props.getOrDefault(MicrosoftSourceConnector.FOLDER_PATH, "/");
        if (grpcTarget.isEmpty()) {
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
        Map<String, Object> offset = context.offsetStorageReader().offset(PARTITION);
        if (offset != null && offset.get("cursor") instanceof String cursor) {
            lastCursor = cursor;
        }
    }

    @Override
    public List<SourceRecord> poll() throws InterruptedException {
        try {
            List<MicrosoftChange> changes = new ArrayList<>();
            String cursor = grpcTarget.isEmpty() ? pollDirect(changes) : pollGrpc(changes);
            lastCursor = cursor;
            Map<String, String> offset = Map.of("cursor", cursor);
            List<SourceRecord> records = new ArrayList<>(changes.size());
            for (MicrosoftChange change : changes) {
                records.add(new SourceRecord(PARTITION, offset, topic,
                        Schema.STRING_SCHEMA, change.getEntity().getEntityId(),
                        Schema.BYTES_SCHEMA, change.toByteArray()));
            }
            return records;
        } catch (InterruptedException e) {
            throw e;
        } catch (Exception e) {
            throw new RuntimeException("microsoft source poll failed", e);
        }
    }

    private String pollGrpc(List<MicrosoftChange> out) {
        ManagedChannelBuilder<?> builder = ManagedChannelBuilder.forTarget(grpcTarget);
        if (plaintext) {
            builder.usePlaintext();
        }
        ManagedChannel channel = builder.build();
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                    MicrosoftServiceGrpc.newBlockingStub(channel);
            SyncRequest.Builder request = SyncRequest.newBuilder()
                    .setIncludeContent(includeContent);
            request.addAllDriveIds(driveIds).setFolderPath(folderPath);
            Iterator<SyncResponse> events = stub.sync(request.build());
            String cursor = lastCursor == null ? Instant.now().toString() : lastCursor;
            while (events.hasNext()) {
                SyncResponse event = events.next();
                if (event.hasChange()) {
                    out.add(event.getChange());
                } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                    cursor = event.getResumeCursor();
                }
            }
            return cursor;
        } finally {
            channel.shutdownNow();
        }
    }

    private String pollDirect(List<MicrosoftChange> out) throws Exception {
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
    public void stop() {
    }

    private static List<String> splitCsv(String raw) {
        if (raw == null || raw.isBlank()) {
            return List.of();
        }
        return java.util.Arrays.stream(raw.split(","))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .toList();
    }

    private static String required(Map<String, String> props, String key) {
        String value = props.get(key);
        if (value == null || value.isBlank()) {
            throw new IllegalArgumentException(key + " is required");
        }
        return value;
    }
}
