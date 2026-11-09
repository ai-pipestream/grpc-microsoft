package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.sync.v1.Connection;
import ai.pipestream.sync.v1.ConnectionKind;
import ai.pipestream.sync.v1.ConnectionServiceGrpc;
import ai.pipestream.sync.v1.GetConnectionRequest;
import ai.pipestream.sync.v1.RecordProbeRequest;
import io.grpc.Status;

import java.io.IOException;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Supplier;

/**
 * Resolves a request's {@code connection_id} to an authorized Graph session: the
 * process-default (environment) connection, a catalog connection read from
 * {@code ConnectionService} (secret included, held in memory only), or an
 * unsaved inline credential set for {@code ProbeConnection}. Catalog and inline
 * sessions use the client-credentials flow against the default config's
 * authority and Graph endpoint, so their auth mode is application.
 */
final class MicrosoftSessions {

    /** Catalog id that means the process-default connection. */
    static final String DEFAULT_CONNECTION_ID = SyncTableMicrosoftChangeSink.DEFAULT_CONNECTION_ID;

    /**
     * One resolved connection.
     *
     * @param connectionId catalog id, {@link #DEFAULT_CONNECTION_ID}, or empty for an
     *        inline probe
     * @param config the connection's crawl scope (site id, drive allowlist)
     * @param files the authorized Graph files API
     */
    record Session(String connectionId, MicrosoftConnectorConfig config, GraphFiles files) {

        GraphClient graph() {
            return files.client();
        }

        GraphPager pager() {
            return new GraphPager(files.client());
        }

        AuthMode authMode() {
            return files.client().authMode();
        }
    }

    private record Cached(long updatedAtSeconds, Session session) {
    }

    private final MicrosoftConnectorConfig defaultConfig;
    private final GraphFiles defaultFiles;
    private final ConnectionServiceGrpc.ConnectionServiceBlockingStub connections;
    private final Map<String, Cached> cache = new ConcurrentHashMap<>();

    MicrosoftSessions(MicrosoftConnectorConfig defaultConfig, GraphFiles defaultFiles,
            ConnectionServiceGrpc.ConnectionServiceBlockingStub connections) {
        this.defaultConfig = Objects.requireNonNull(defaultConfig, "defaultConfig");
        this.defaultFiles = Objects.requireNonNull(defaultFiles, "defaultFiles");
        this.connections = connections;
    }

    /**
     * The session for {@code connectionId}.
     *
     * @param connectionId empty or {@code default} = the environment connection
     * @return the session
     */
    Session resolve(String connectionId) {
        if (connectionId == null || connectionId.isBlank()
                || connectionId.equals(DEFAULT_CONNECTION_ID)) {
            return new Session(DEFAULT_CONNECTION_ID, defaultConfig, defaultFiles);
        }
        if (connections == null) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("connection_id " + connectionId
                            + " needs ConnectionService, which is not configured")
                    .asRuntimeException();
        }
        Connection row = connections.getConnection(GetConnectionRequest.newBuilder()
                .setConnectionId(connectionId)
                .setIncludeSecret(true)
                .build()).getConnection();
        if (row.getKind() != ConnectionKind.CONNECTION_KIND_UNSPECIFIED
                && row.getKind() != ConnectionKind.CONNECTION_KIND_MICROSOFT) {
            throw Status.INVALID_ARGUMENT
                    .withDescription("connection " + connectionId + " is not a Microsoft connection")
                    .asRuntimeException();
        }
        if (row.getTenantId().isBlank() || row.getClientId().isBlank()
                || row.getClientSecret().isBlank()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("connection " + connectionId + " is missing credentials")
                    .asRuntimeException();
        }
        long updated = row.getUpdatedAt().getSeconds();
        Cached cached = cache.get(connectionId);
        if (cached != null && cached.updatedAtSeconds() == updated) {
            return cached.session();
        }
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId(row.getTenantId())
                .clientId(row.getClientId())
                .clientSecret(row.getClientSecret())
                .siteId(row.getSiteId())
                .driveIds(row.getDriveIdsList())
                .graphBaseUrl(defaultConfig.graphBaseUrl())
                .authority(defaultConfig.authority())
                .build();
        Session session = new Session(row.getConnectionId(), config, clientCredentials(config));
        cache.put(connectionId, new Cached(updated, session));
        return session;
    }

    /**
     * A throwaway session for an unsaved credential set.
     *
     * @param tenantId Entra tenant
     * @param clientId application id
     * @param clientSecret client secret; never logged
     * @return the session
     */
    Session inline(String tenantId, String clientId, String clientSecret) {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId(tenantId)
                .clientId(clientId)
                .clientSecret(clientSecret)
                .graphBaseUrl(defaultConfig.graphBaseUrl())
                .authority(defaultConfig.authority())
                .build();
        return new Session("", config, clientCredentials(config));
    }

    /**
     * Records a probe result on the catalog row; best effort.
     *
     * @param connectionId catalog id; blank or default skips
     * @param ok whether the probe worked
     * @param errorMessage failure detail (already scrubbed of secrets)
     */
    void recordProbe(String connectionId, boolean ok, String errorMessage) {
        if (connections == null || connectionId == null || connectionId.isBlank()
                || connectionId.equals(DEFAULT_CONNECTION_ID)) {
            return;
        }
        try {
            RecordProbeRequest.Builder record = RecordProbeRequest.newBuilder()
                    .setConnectionId(connectionId)
                    .setOk(ok);
            if (!ok && errorMessage != null) {
                record.setErrorMessage(errorMessage);
            }
            connections.recordProbe(record.build());
        } catch (RuntimeException ignored) {
            // catalog status is best-effort
        }
    }

    private static GraphFiles clientCredentials(MicrosoftConnectorConfig config) {
        GraphAuth auth = new GraphAuth(config.authConfig());
        return new GraphFiles(new GraphClient(config.graphBaseUrl(), tokenSupplier(auth),
                AuthMode.AUTH_MODE_APPLICATION));
    }

    /**
     * A lazily refreshed client-credentials token. A token failure surfaces as an
     * {@link IllegalStateException} whose message never carries the secret.
     *
     * @param auth the Entra helper
     * @return the bearer supplier
     */
    static Supplier<String> tokenSupplier(GraphAuth auth) {
        Object lock = new Object();
        GraphAuth.Token[] holder = new GraphAuth.Token[1];
        return () -> {
            synchronized (lock) {
                if (holder[0] == null || holder[0].expired()) {
                    try {
                        holder[0] = auth.clientCredentials();
                    } catch (IOException e) {
                        throw new IllegalStateException("token request failed: " + e.getMessage(), e);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                        throw new IllegalStateException("token request interrupted", e);
                    }
                }
                return holder[0].accessToken();
            }
        };
    }
}
