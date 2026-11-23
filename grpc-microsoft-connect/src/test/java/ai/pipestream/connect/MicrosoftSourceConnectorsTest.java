package ai.pipestream.connect;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import com.google.protobuf.Timestamp;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpServer;
import io.grpc.Server;
import io.grpc.stub.StreamObserver;
import io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder;
import org.apache.kafka.connect.source.SourceRecord;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.concurrent.ConcurrentLinkedQueue;
import java.util.concurrent.CopyOnWriteArrayList;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class MicrosoftSourceConnectorsTest {

    private Server server;
    private HttpServer graph;

    @AfterEach
    void stop() {
        if (server != null) {
            server.shutdownNow();
        }
        if (graph != null) {
            graph.stop(0);
        }
    }

    @Test
    void microsoftConnectorMetadata() {
        MicrosoftSourceConnector connector = new MicrosoftSourceConnector();
        connector.start(Map.of("topic", "microsoft.changes"));
        assertThat(connector.version()).isEqualTo(ConnectVersion.VALUE);
        assertThat(connector.taskClass()).isEqualTo(MicrosoftSourceTask.class);
        assertThat(connector.taskConfigs(3)).hasSize(1);
        assertThat(connector.config().configKeys()).containsKeys(MicrosoftSourceConnector.TOPIC,
                MicrosoftSourceConnector.TENANT_ID, MicrosoftSourceConnector.DRIVE_IDS);
        connector.stop();
    }

    @Test
    void microsoftGrpcPollEmitsProtobufBytes() throws Exception {
        server = NettyServerBuilder.forAddress(new InetSocketAddress("127.0.0.1", 0))
                .addService(new MicrosoftServiceGrpc.MicrosoftServiceImplBase() {
                    @Override
                    public void sync(ai.pipestream.microsoft.v1.SyncRequest request,
                            StreamObserver<ai.pipestream.microsoft.v1.SyncResponse> observer) {
                        observer.onNext(ai.pipestream.microsoft.v1.SyncResponse.newBuilder()
                                .setChange(MicrosoftChange.newBuilder()
                                        .setChangeId("c1")
                                        .setOperation(
                                                ai.pipestream.microsoft.v1.ChangeOperation.CHANGE_OPERATION_UPSERT)
                                        .setEntity(MicrosoftEntity.newBuilder()
                                                .setEntityId("file-1")
                                                .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                                                .setDriveItem(DriveItem.newBuilder()
                                                        .setId("file-1")
                                                        .setName("notes.txt"))))
                                .build());
                        observer.onNext(ai.pipestream.microsoft.v1.SyncResponse.newBuilder()
                                .setResumeCursor("cursor-9")
                                .build());
                        observer.onCompleted();
                    }
                })
                .build()
                .start();

        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.empty());
        task.start(Map.of(
                MicrosoftSourceConnector.TOPIC, "ms.changes",
                MicrosoftSourceConnector.GRPC_TARGET, "127.0.0.1:" + server.getPort(),
                MicrosoftSourceConnector.GRPC_PLAINTEXT, "true"));
        List<SourceRecord> records = task.poll();
        task.stop();

        assertThat(records).hasSize(1);
        assertThat(records.get(0).topic()).isEqualTo("ms.changes");
        MicrosoftChange parsed = MicrosoftChange.parseFrom((byte[]) records.get(0).value());
        assertThat(parsed.getEntity().getEntityId()).isEqualTo("file-1");
        assertThat(records.get(0).sourceOffset().get("cursor")).isEqualTo("cursor-9");
    }

    @Test
    void grpcPollSendsTheStoredCursorAndThenItsOwn() throws Exception {
        RecordingSync sync = new RecordingSync("cursor-9", "cursor-10");
        server = start(sync);

        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.withOffset(Map.of("cursor", "cursor-8")));
        task.start(grpcProps());
        List<SourceRecord> first = task.poll();
        List<SourceRecord> second = task.poll();
        task.stop();

        // Each pass is a delta pass from the cursor before it, so deletes
        // arrive and the drive is not republished.
        assertThat(sync.requests).extracting(r -> r.getSinceCursor())
                .containsExactly("cursor-8", "cursor-9");
        assertThat(first.get(0).sourceOffset().get("cursor")).isEqualTo("cursor-9");
        assertThat(second.get(0).sourceOffset().get("cursor")).isEqualTo("cursor-10");
    }

    @Test
    void grpcPollWithNoOffsetAsksForAFullEnumeration() throws Exception {
        RecordingSync sync = new RecordingSync("cursor-9");
        server = start(sync);

        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.empty());
        task.start(grpcProps());
        task.poll();
        task.stop();

        assertThat(sync.requests).extracting(r -> r.getSinceCursor()).containsExactly("");
    }

    @Test
    void aWallClockOffsetFromAnOlderVersionIsNotSentAsACursor() throws Exception {
        RecordingSync sync = new RecordingSync("cursor-9");
        server = start(sync);

        // Older versions stored the poll's wall-clock time as the offset.
        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.withOffset(
                Map.of("cursor", "2026-11-01T08:15:30.123456789Z")));
        task.start(grpcProps());
        List<SourceRecord> records = task.poll();
        task.stop();

        assertThat(sync.requests).extracting(r -> r.getSinceCursor()).containsExactly("");
        assertThat(records.get(0).sourceOffset().get("cursor")).isEqualTo("cursor-9");
    }

    @Test
    void grpcPollFailsWhenTheStreamEndsWithoutItsResumeCursor() throws Exception {
        RecordingSync sync = new RecordingSync();
        server = start(sync);

        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.withOffset(Map.of("cursor", "cursor-8")));
        task.start(grpcProps());

        // A truncated pass commits nothing: no records, no new offset.
        assertThatThrownBy(task::poll)
                .hasRootCauseMessage("Sync ended without its resume cursor");
        task.stop();
    }

    @Test
    void directPollResumesFromItsOffsetAndPublishesTheDelete() throws Exception {
        graph = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        String base = "http://127.0.0.1:" + graph.getAddress().getPort();
        List<String> deltaQueries = new CopyOnWriteArrayList<>();
        graph.createContext("/tenant-1/oauth2/v2.0/token", exchange -> reply(exchange, 200,
                "{\"access_token\":\"t\",\"token_type\":\"Bearer\",\"expires_in\":3600}"));
        graph.createContext("/drives/drive-1", exchange -> {
            String path = exchange.getRequestURI().getPath();
            String query = exchange.getRequestURI().getRawQuery();
            if (path.equals("/drives/drive-1")) {
                reply(exchange, 200, "{\"id\":\"drive-1\",\"name\":\"Docs\","
                        + "\"driveType\":\"business\"}");
            } else if (path.equals("/drives/drive-1/root/delta")) {
                deltaQueries.add(query == null ? "" : query);
                if ("token=d1".equals(query)) {
                    reply(exchange, 200, "{\"value\":[{\"id\":\"file-1\",\"file\":{},"
                            + "\"deleted\":{},\"parentReference\":{\"id\":\"root\","
                            + "\"driveId\":\"drive-1\"}}],\"@odata.deltaLink\":\""
                            + base + "/drives/drive-1/root/delta?token=d2\"}");
                } else {
                    reply(exchange, 200, "{\"value\":[{\"id\":\"file-1\","
                            + "\"name\":\"notes.txt\",\"size\":12,"
                            + "\"file\":{\"mimeType\":\"text/plain\"},"
                            + "\"parentReference\":{\"id\":\"root\",\"driveId\":\"drive-1\"}}],"
                            + "\"@odata.deltaLink\":\"" + base
                            + "/drives/drive-1/root/delta?token=d1\"}");
                }
            } else {
                reply(exchange, 404, "{\"error\":{\"code\":\"notFound\",\"message\":\"none\"}}");
            }
        });
        graph.start();

        MicrosoftSourceTask task = new MicrosoftSourceTask();
        task.initialize(ConnectTestContexts.empty());
        task.start(Map.of(
                MicrosoftSourceConnector.TOPIC, "ms.changes",
                MicrosoftSourceConnector.TENANT_ID, "tenant-1",
                MicrosoftSourceConnector.CLIENT_ID, "client-1",
                MicrosoftSourceConnector.CLIENT_SECRET, "secret-1",
                MicrosoftSourceConnector.DRIVE_IDS, "drive-1",
                MicrosoftSourceConnector.GRAPH_BASE_URL, base,
                MicrosoftSourceConnector.AUTHORITY, base));
        List<SourceRecord> first = task.poll();
        List<SourceRecord> second = task.poll();
        task.stop();

        // The first pass enumerates; its offset is the encoded SyncCursor
        // holding the drive's delta link, not a wall-clock string.
        assertThat(changes(first)).extracting(c -> c.getEntity().getEntityId())
                .contains("file-1");
        String offset = (String) first.get(0).sourceOffset().get("cursor");
        assertThat(SyncCursor.parseFrom(Base64.getUrlDecoder().decode(offset))
                .getDriveDeltaLinksMap())
                .containsEntry("drive-1", base + "/drives/drive-1/root/delta?token=d1");

        // The second pass reads the delta link, so the delete is published.
        assertThat(deltaQueries.get(deltaQueries.size() - 1)).isEqualTo("token=d1");
        assertThat(changes(second))
                .filteredOn(c -> c.getEntity().getEntityId().equals("file-1"))
                .extracting(MicrosoftChange::getOperation)
                .containsExactly(ChangeOperation.CHANGE_OPERATION_DELETE);
        String next = (String) second.get(0).sourceOffset().get("cursor");
        assertThat(SyncCursor.parseFrom(Base64.getUrlDecoder().decode(next))
                .getDriveDeltaLinksMap())
                .containsEntry("drive-1", base + "/drives/drive-1/root/delta?token=d2");
    }

    private Map<String, String> grpcProps() {
        return Map.of(
                MicrosoftSourceConnector.TOPIC, "ms.changes",
                MicrosoftSourceConnector.GRPC_TARGET, "127.0.0.1:" + server.getPort(),
                MicrosoftSourceConnector.GRPC_PLAINTEXT, "true");
    }

    private static Server start(RecordingSync sync) throws IOException {
        return NettyServerBuilder.forAddress(new InetSocketAddress("127.0.0.1", 0))
                .addService(sync)
                .build()
                .start();
    }

    private static List<MicrosoftChange> changes(List<SourceRecord> records) throws Exception {
        List<MicrosoftChange> out = new ArrayList<>();
        for (SourceRecord record : records) {
            out.add(MicrosoftChange.parseFrom((byte[]) record.value()));
        }
        return out;
    }

    private static void reply(HttpExchange exchange, int status, String json) throws IOException {
        try (exchange) {
            exchange.getRequestBody().readAllBytes();
            byte[] body = json.getBytes(StandardCharsets.UTF_8);
            exchange.getResponseHeaders().add("content-type", "application/json");
            exchange.sendResponseHeaders(status, body.length);
            exchange.getResponseBody().write(body);
        }
    }

    /**
     * A MicrosoftService whose every Sync sends one change and then the
     * next of {@code cursors}; with none left it ends the stream without a
     * resume cursor.
     */
    private static final class RecordingSync extends MicrosoftServiceGrpc.MicrosoftServiceImplBase {

        final List<SyncRequest> requests = new CopyOnWriteArrayList<>();
        private final Queue<String> cursors;

        RecordingSync(String... cursors) {
            this.cursors = new ConcurrentLinkedQueue<>(List.of(cursors));
        }

        @Override
        public void sync(SyncRequest request, StreamObserver<SyncResponse> observer) {
            requests.add(request);
            observer.onNext(SyncResponse.newBuilder()
                    .setChange(MicrosoftChange.newBuilder()
                            .setChangeId("c" + requests.size())
                            .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                            .setEntity(MicrosoftEntity.newBuilder()
                                    .setEntityId("file-1")
                                    .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                                    .setDriveItem(DriveItem.newBuilder()
                                            .setId("file-1")
                                            .setName("notes.txt"))))
                    .build());
            String cursor = cursors.poll();
            if (cursor != null) {
                observer.onNext(SyncResponse.newBuilder().setResumeCursor(cursor).build());
            }
            observer.onCompleted();
        }
    }
}
