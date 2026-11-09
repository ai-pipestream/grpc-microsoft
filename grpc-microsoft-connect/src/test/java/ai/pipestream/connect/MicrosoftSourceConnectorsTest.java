package ai.pipestream.connect;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import com.google.protobuf.Timestamp;
import io.grpc.Server;
import io.grpc.stub.StreamObserver;
import io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder;
import org.apache.kafka.connect.source.SourceRecord;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.net.InetSocketAddress;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftSourceConnectorsTest {

    private Server server;

    @AfterEach
    void stop() {
        if (server != null) {
            server.shutdownNow();
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
}
