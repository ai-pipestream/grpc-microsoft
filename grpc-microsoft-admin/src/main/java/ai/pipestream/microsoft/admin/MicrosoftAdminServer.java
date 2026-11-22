package ai.pipestream.microsoft.admin;

import ai.pipestream.microsoft.GraphAuth;
import ai.pipestream.microsoft.GraphClient;
import ai.pipestream.microsoft.MicrosoftServer;
import ai.pipestream.microsoft.v1.AuthMode;
import io.grpc.Server;
import java.util.concurrent.TimeUnit;

/** Standalone server of {@code MicrosoftAdminService}, signed in as the admin identity. */
public final class MicrosoftAdminServer {

    /** Environment variable for the gRPC listen port. */
    public static final String ENV_GRPC_PORT = "MICROSOFT_ADMIN_GRPC_PORT";
    /** Default gRPC listen port when {@link #ENV_GRPC_PORT} is unset. */
    public static final int DEFAULT_GRPC_PORT = 9097;

    private static final System.Logger LOG = System.getLogger(MicrosoftAdminServer.class.getName());

    private MicrosoftAdminServer() {
    }

    /**
     * Starts the server from the process environment and blocks until shutdown.
     *
     * @param args unused
     * @throws Exception if the server cannot start or is interrupted
     */
    public static void main(String[] args) throws Exception {
        MicrosoftAdminConfig config = MicrosoftAdminConfig.fromEnvironment();
        Server server = MicrosoftServer.startNetty(service(config), port(System.getenv(ENV_GRPC_PORT)));
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
        }, "microsoft-admin-shutdown"));
        LOG.log(System.Logger.Level.INFO, "microsoft-admin listening on port {0} (graph {1})",
                server.getPort(), config.graphBaseUrl());
        server.awaitTermination();
    }

    /**
     * The service for a config.
     *
     * @param config the admin identity and endpoints
     * @return the service, signed in on first use
     */
    public static MicrosoftAdminGrpcService service(MicrosoftAdminConfig config) {
        GraphAuth auth = new GraphAuth(config.authConfig());
        return new MicrosoftAdminGrpcService(new GraphClient(config.graphBaseUrl(),
                auth.tokens(GraphAuth.GRAPH_SCOPE), AuthMode.AUTH_MODE_APPLICATION));
    }

    private static int port(String value) {
        if (value == null || value.isBlank()) {
            return DEFAULT_GRPC_PORT;
        }
        try {
            return Integer.parseInt(value.trim());
        } catch (NumberFormatException e) {
            return DEFAULT_GRPC_PORT;
        }
    }
}
