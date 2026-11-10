package ai.pipestream.microsoft.connector;

import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;
import io.grpc.Server;
import io.grpc.health.v1.HealthCheckRequest;
import io.grpc.health.v1.HealthCheckResponse;
import io.grpc.health.v1.HealthGrpc;
import microsoft.graph.connectors.contracts.grpc.ConnectorInfoServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.GetBasicConnectorInfoRequest;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class ConnectorServerTest {

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
    void bindsAndServesInfoAndHealth() throws Exception {
        server = ConnectorServer.startNetty(0);
        channel = ManagedChannelBuilder.forAddress("127.0.0.1", server.getPort())
                .usePlaintext()
                .build();

        assertThat(HealthGrpc.newBlockingStub(channel)
                .check(HealthCheckRequest.getDefaultInstance())
                .getStatus()).isEqualTo(HealthCheckResponse.ServingStatus.SERVING);
        assertThat(ConnectorInfoServiceGrpc.newBlockingStub(channel)
                .getBasicConnectorInfo(GetBasicConnectorInfoRequest.getDefaultInstance())
                .getConnectorId()).isEqualTo(ConnectorId.VALUE);
    }
}
