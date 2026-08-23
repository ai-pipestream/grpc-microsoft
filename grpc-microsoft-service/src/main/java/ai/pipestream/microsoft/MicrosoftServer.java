package ai.pipestream.microsoft;

import io.grpc.BindableService;
import io.grpc.Server;
import io.grpc.health.v1.HealthCheckResponse;
import io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder;
import io.grpc.protobuf.services.ProtoReflectionService;
import io.grpc.protobuf.services.ProtoReflectionServiceV1;
import io.grpc.services.HealthStatusManager;

import java.io.IOException;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

/**
 * Standalone Microsoft Graph gRPC proxy. Client-credentials token is
 * refreshed in memory; handlers run on virtual threads.
 */
public final class MicrosoftServer {

    /** Environment variable for the gRPC listen port. */
    public static final String ENV_GRPC_PORT = "MICROSOFT_GRPC_PORT";
    /** Environment variable for the inline file byte cap. */
    public static final String ENV_ATTACHMENT_MAX_BYTES = "MICROSOFT_ATTACHMENT_MAX_BYTES";
    /** Default gRPC listen port when {@link #ENV_GRPC_PORT} is unset. */
    public static final int DEFAULT_GRPC_PORT = 9096;

    private static final System.Logger LOG = System.getLogger(MicrosoftServer.class.getName());

    private MicrosoftServer() {
    }

    /**
     * Starts the proxy from the process environment and blocks until shutdown.
     *
     * @param args unused
     * @throws Exception if the server cannot start or is interrupted
     */
    public static void main(String[] args) throws Exception {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment();
        GraphAuth auth = new GraphAuth(config.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(auth.clientCredentials());
        GraphClient client = new GraphClient(config.graphBaseUrl(), () -> {
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
        });
        MicrosoftChangeSink downstream = null;
        java.util.List<AutoCloseable> closables = new java.util.ArrayList<>();
        java.util.List<MicrosoftChangeSink> sinks = new java.util.ArrayList<>();
        GraphFiles files = new GraphFiles(client);
        if (SyncTableMicrosoftChangeSink.enabled()) {
            SyncTableMicrosoftChangeSink syncTable = SyncTableMicrosoftChangeSink.fromEnvironment();
            sinks.add(syncTable);
            closables.add(syncTable);
            LOG.log(System.Logger.Level.INFO, "microsoft-proxy sync-table sink active on {0}",
                    System.getenv(SyncTableMicrosoftChangeSink.ENV_TARGET));
        }
        if (OutputMicrosoftChangeSink.enabled()) {
            OutputMicrosoftChangeSink output = OutputMicrosoftChangeSink.fromEnvironment();
            sinks.add(output);
            closables.add(output);
        }
        if (OkfMicrosoftChangeSink.enabled() && SharePointOkfPublisher.enabled()) {
            sinks.add(OkfMicrosoftChangeSink.fromEnvironment(files));
            LOG.log(System.Logger.Level.INFO,
                    "microsoft-proxy SharePoint OKF upload active spoDrive={0}",
                    System.getenv(SharePointOkfPublisher.ENV_DRIVE_ID));
        } else if (!OutputMicrosoftChangeSink.enabled() && OkfMicrosoftChangeSink.enabled()) {
            sinks.add(OkfMicrosoftChangeSink.fromEnvironment(files));
            LOG.log(System.Logger.Level.INFO,
                    "microsoft-proxy OKF sink active dir={0} spoDrive={1}",
                    System.getenv(ai.pipestream.okf.OkfOutput.ENV_DIR),
                    System.getenv(SharePointOkfPublisher.ENV_DRIVE_ID));
        }
        if (!sinks.isEmpty()) {
            downstream = sinks.size() == 1 ? sinks.get(0)
                    : new CompositeMicrosoftChangeSink(sinks);
        }
        MicrosoftGrpcService service = new MicrosoftGrpcService(config, files,
                parseLong(System.getenv(ENV_ATTACHMENT_MAX_BYTES),
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES),
                downstream);
        Server server = startNetty(service, parseInt(System.getenv(ENV_GRPC_PORT),
                DEFAULT_GRPC_PORT));
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            server.shutdown();
            try {
                if (!server.awaitTermination(10, TimeUnit.SECONDS)) {
                    server.shutdownNow();
                }
                for (AutoCloseable closable : closables) {
                    closable.close();
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                server.shutdownNow();
            } catch (Exception e) {
                LOG.log(System.Logger.Level.WARNING, "sink close failed: {0}", e.toString());
            }
        }, "microsoft-proxy-shutdown"));
        LOG.log(System.Logger.Level.INFO,
                "microsoft-proxy listening on port {0} (graph {1}, site {2})",
                server.getPort(), config.graphBaseUrl(),
                config.siteId().isBlank() ? "me-drive" : config.siteId());
        server.awaitTermination();
    }

    /**
     * Binds {@code service} on Netty with reflection, health, and a virtual-thread executor.
     *
     * @param service the gRPC service to serve
     * @param port the listen port
     * @return the started server
     * @throws IOException if the port cannot be bound
     */
    public static Server startNetty(BindableService service, int port) throws IOException {
        HealthStatusManager health = new HealthStatusManager();
        Server server = NettyServerBuilder.forPort(port)
                .executor(Executors.newVirtualThreadPerTaskExecutor())
                .addService(service)
                .addService(health.getHealthService())
                .addService(ProtoReflectionService.newInstance())
                .addService(ProtoReflectionServiceV1.newInstance())
                .build().start();
        health.setStatus("", HealthCheckResponse.ServingStatus.SERVING);
        return server;
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

    private static long parseLong(String value, long fallback) {
        if (value == null || value.isBlank()) {
            return fallback;
        }
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }
}
