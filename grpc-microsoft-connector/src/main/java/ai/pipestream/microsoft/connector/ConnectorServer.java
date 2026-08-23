package ai.pipestream.microsoft.connector;

import io.grpc.Server;
import io.grpc.health.v1.HealthCheckResponse;
import io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder;
import io.grpc.protobuf.services.ProtoReflectionService;
import io.grpc.protobuf.services.ProtoReflectionServiceV1;
import io.grpc.services.HealthStatusManager;

import java.io.IOException;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

/**
 * GCA-facing Copilot connector process. Listens on port 30303 by default
 * (the Graph Connector Agent template port) and implements the four SDK
 * services. Microsoft Graph traffic goes to {@code MicrosoftService}.
 *
 * <p>Environment:</p>
 * <ul>
 *   <li>{@code CONNECTOR_GRPC_PORT}: listen port, default 30303</li>
 *   <li>{@code MICROSOFT_GRPC_TARGET}: default MicrosoftService
 *   {@code host:port} when custom configuration omits {@code target}</li>
 *   <li>{@code MICROSOFT_GRPC_PLAINTEXT}: default true</li>
 * </ul>
 */
public final class ConnectorServer {

    /** Environment variable for the listen port. */
    public static final String ENV_GRPC_PORT = "CONNECTOR_GRPC_PORT";
    /** Default GCA template listen port. */
    public static final int DEFAULT_GRPC_PORT = 30303;

    private static final System.Logger LOG = System.getLogger(ConnectorServer.class.getName());

    private ConnectorServer() {
    }

    /**
     * Starts the connector process and blocks until shutdown.
     *
     * @param args unused
     * @throws Exception if the server fails to start or is interrupted
     */
    public static void main(String[] args) throws Exception {
        Server server = startNetty(parseInt(System.getenv(ENV_GRPC_PORT), DEFAULT_GRPC_PORT));
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            server.shutdown();
            try {
                if (!server.awaitTermination(10, TimeUnit.SECONDS)) {
                    server.shutdownNow();
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                server.shutdownNow();
            }
        }, "connector-shutdown"));
        LOG.log(System.Logger.Level.INFO,
                "copilot-connector {0} listening on port {1} (MicrosoftService default {2})",
                ConnectorId.VALUE, server.getPort(),
                firstNonBlank(System.getenv(ConnectorCustomConfig.ENV_TARGET),
                        ConnectorCustomConfig.DEFAULT_TARGET));
        server.awaitTermination();
    }

    /**
     * Starts a Netty gRPC server on {@code port} with the four GCA services,
     * health, and reflection.
     *
     * @param port listen port
     * @return the started server
     * @throws IOException if the port cannot be bound
     */
    public static Server startNetty(int port) throws IOException {
        HealthStatusManager health = new HealthStatusManager();
        Server server = NettyServerBuilder.forPort(port)
                .executor(Executors.newVirtualThreadPerTaskExecutor())
                .addService(new ConnectorInfoServiceImpl())
                .addService(new ConnectionManagementServiceImpl())
                .addService(new ConnectorCrawlerServiceImpl())
                .addService(new ConnectorOAuthServiceImpl())
                .addService(health.getHealthService())
                .addService(ProtoReflectionService.newInstance())
                .addService(ProtoReflectionServiceV1.newInstance())
                .build()
                .start();
        health.setStatus("", HealthCheckResponse.ServingStatus.SERVING);
        return server;
    }

    private static String firstNonBlank(String value, String fallback) {
        return value == null || value.isBlank() ? fallback : value;
    }

    private static int parseInt(String value, int fallback) {
        if (value == null || value.isBlank()) {
            return fallback;
        }
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }
}
