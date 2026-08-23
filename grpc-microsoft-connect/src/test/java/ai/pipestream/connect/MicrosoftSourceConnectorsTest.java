package ai.pipestream.connect;

import ai.pipestream.confluence.v1.ChangeOperation;
import ai.pipestream.confluence.v1.ConfluenceChange;
import ai.pipestream.confluence.v1.ConfluenceEntity;
import ai.pipestream.confluence.v1.ConfluenceServiceGrpc;
import ai.pipestream.confluence.v1.Page;
import ai.pipestream.confluence.v1.SyncRequest;
import ai.pipestream.confluence.v1.SyncResponse;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import com.google.protobuf.Timestamp;
import io.grpc.Server;
import io.grpc.stub.StreamObserver;
import io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder;
import org.apache.kafka.connect.source.SourceRecord;
import org.apache.kafka.connect.source.SourceTaskContext;
import org.apache.kafka.connect.storage.OffsetStorageReader;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.net.InetSocketAddress;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class SourceConnectorsTest {

    private Server server;

    @AfterEach
    void stop() {
        if (server != null) {
            server.shutdownNow();
        }
    }

    @Test
    void confluenceConnectorMetadata() {
        ConfluenceSourceConnector connector = new ConfluenceSourceConnector();
        connector.start(Map.of("topic", "confluence.changes"));
        assertThat(connector.version()).isEqualTo(ConnectVersion.VALUE);
        assertThat(connector.taskClass()).isEqualTo(ConfluenceSourceTask.class);
        assertThat(connector.taskConfigs(3)).hasSize(1);
        assertThat(connector.config().configKeys()).containsKey(ConfluenceSourceConnector.TOPIC);
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
        task.initialize(emptyContext());
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
        assertThat(records.get(0).sourceOffset()).containsEntry("cursor", "cursor-9");
    }

    @Test
    void confluenceGrpcPollEmitsProtobufBytes() throws Exception {
        server = NettyServerBuilder.forAddress(new InetSocketAddress("127.0.0.1", 0))
                .addService(new ConfluenceServiceGrpc.ConfluenceServiceImplBase() {
                    @Override
                    public void sync(SyncRequest request, StreamObserver<SyncResponse> observer) {
                        observer.onNext(SyncResponse.newBuilder()
                                .setChange(ConfluenceChange.newBuilder()
                                        .setChangeId("c1")
                                        .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                                        .setEntity(ConfluenceEntity.newBuilder()
                                                .setEntityId("page-1")
                                                .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                                                .setPage(Page.newBuilder()
                                                        .setId("page-1")
                                                        .setSpaceId("100")
                                                        .setTitle("Design"))))
                                .build());
                        observer.onNext(SyncResponse.newBuilder().setResumeCursor("2024-03-02T00:00:00Z")
                                .build());
                        observer.onCompleted();
                    }
                })
                .build()
                .start();

        ConfluenceSourceTask task = new ConfluenceSourceTask();
        task.initialize(emptyContext());
        task.start(Map.of(
                ConfluenceSourceConnector.TOPIC, "conf.changes",
                ConfluenceSourceConnector.GRPC_TARGET, "127.0.0.1:" + server.getPort(),
                ConfluenceSourceConnector.GRPC_PLAINTEXT, "true"));
        List<SourceRecord> records = task.poll();
        task.stop();

        assertThat(records).hasSize(1);
        ConfluenceChange parsed = ConfluenceChange.parseFrom((byte[]) records.get(0).value());
        assertThat(parsed.getEntity().getPage().getTitle()).isEqualTo("Design");
    }

    private static SourceTaskContext emptyContext() {
        return new SourceTaskContext() {
            @Override
            public OffsetStorageReader offsetStorageReader() {
                return new OffsetStorageReader() {
                    @Override
                    public <T> Map<Map<String, T>, Map<String, Object>> offsets(
                            Collection<Map<String, T>> partitions) {
                        return Map.of();
                    }

                    @Override
                    public <T> Map<String, Object> offset(Map<String, T> partition) {
                        return null;
                    }
                };
            }

            @Override
            public Map<String, String> configs() {
                return Map.of();
            }
        };
    }
}
