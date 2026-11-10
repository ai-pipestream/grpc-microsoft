package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;
import io.grpc.Server;
import io.grpc.health.v1.HealthCheckRequest;
import io.grpc.health.v1.HealthCheckResponse;
import io.grpc.health.v1.HealthGrpc;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftServerTest {

    private Server server;
    private ManagedChannel channel;

    @AfterEach
    void stop() {
        if (channel != null) {
            channel.shutdownNow();
        }
        if (server != null) {
            server.shutdownNow();
        }
    }

    @Test
    void bindsAndServesHealth() throws Exception {
        server = MicrosoftServer.startNetty(
                new MicrosoftServiceGrpc.MicrosoftServiceImplBase() {
                }, 0);
        channel = ManagedChannelBuilder.forAddress("127.0.0.1", server.getPort())
                .usePlaintext()
                .build();
        assertThat(server.getPort()).isPositive();
        assertThat(HealthGrpc.newBlockingStub(channel)
                .check(HealthCheckRequest.getDefaultInstance())
                .getStatus()).isEqualTo(HealthCheckResponse.ServingStatus.SERVING);
    }
}
