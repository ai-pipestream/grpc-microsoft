package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.sync.v1.Asset;
import ai.pipestream.sync.v1.AssetPhase;
import ai.pipestream.sync.v1.AssetSyncStatus;
import ai.pipestream.sync.v1.Checkpoint;
import ai.pipestream.sync.v1.DeleteAssetRequest;
import ai.pipestream.sync.v1.PutCheckpointRequest;
import ai.pipestream.sync.v1.ReconcileRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import ai.pipestream.sync.v1.UpsertAssetRequest;
import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;

import java.util.Locale;
import java.util.Objects;
import java.util.concurrent.TimeUnit;

/** Writes Microsoft Graph changes into {@code SyncTableService}. */
public final class SyncTableMicrosoftChangeSink implements MicrosoftChangeSink, AutoCloseable {

    /** Asset {@code source} written for every Microsoft Graph row. */
    public static final String SOURCE = "microsoft";
    /** Catalog id used when the caller does not name a connection. */
    public static final String DEFAULT_CONNECTION_ID = "default";
    /** Environment variable for the {@code SyncTableService} target. */
    public static final String ENV_TARGET = "SYNC_TABLE_TARGET";
    /** Environment variable; {@code false} disables plaintext on the channel. */
    public static final String ENV_PLAINTEXT = "SYNC_TABLE_PLAINTEXT";

    private static final System.Logger LOG =
            System.getLogger(SyncTableMicrosoftChangeSink.class.getName());

    private final ManagedChannel channel;
    private final SyncTableServiceGrpc.SyncTableServiceBlockingStub stub;
    private final String connectionId;

    /**
     * Creates a sink over an existing channel (takes ownership for {@link #close()}).
     * Rows are stamped {@link #DEFAULT_CONNECTION_ID}.
     *
     * @param channel the gRPC channel to {@code SyncTableService}
     */
    public SyncTableMicrosoftChangeSink(ManagedChannel channel) {
        this(channel, DEFAULT_CONNECTION_ID);
    }

    /**
     * Creates a sink that stamps {@code connectionId} on every row.
     *
     * @param channel the gRPC channel to {@code SyncTableService}
     * @param connectionId catalog connection
     */
    public SyncTableMicrosoftChangeSink(ManagedChannel channel, String connectionId) {
        this.channel = Objects.requireNonNull(channel, "channel");
        this.stub = SyncTableServiceGrpc.newBlockingStub(channel);
        this.connectionId = connectionId == null || connectionId.isBlank()
                ? DEFAULT_CONNECTION_ID : connectionId;
    }

    /**
     * Whether sync-table publishing is configured in the process environment.
     *
     * @return {@code true} when {@link #ENV_TARGET} is set
     */
    public static boolean enabled() {
        String target = System.getenv(ENV_TARGET);
        return target != null && !target.isBlank();
    }

    /**
     * Builds a sink from {@code SYNC_TABLE_*} environment variables.
     *
     * @return a sink connected to the configured target
     */
    public static SyncTableMicrosoftChangeSink fromEnvironment() {
        String target = System.getenv(ENV_TARGET);
        if (target == null || target.isBlank()) {
            throw new IllegalStateException(ENV_TARGET + " is required");
        }
        boolean plaintext = !"false".equalsIgnoreCase(
                Objects.toString(System.getenv(ENV_PLAINTEXT), "true"));
        ManagedChannelBuilder<?> builder = ManagedChannelBuilder.forTarget(target.trim());
        if (plaintext) {
            builder.usePlaintext();
        }
        return new SyncTableMicrosoftChangeSink(builder.build());
    }

    @Override
    public void emit(MicrosoftChange change) {
        try {
            if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
                stub.deleteAsset(DeleteAssetRequest.newBuilder()
                        .setAssetId(assetId(connectionId, change.getEntity()))
                        .setRunId(change.getCursor())
                        .setCursor(change.getCursor())
                        .build());
                return;
            }
            stub.upsertAsset(UpsertAssetRequest.newBuilder()
                    .setAsset(toAsset(connectionId, change)).build());
        } catch (RuntimeException e) {
            LOG.log(System.Logger.Level.WARNING, "sync-table emit failed: {0}", e.toString());
        }
    }

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        try {
            stub.putCheckpoint(PutCheckpointRequest.newBuilder()
                    .setCheckpoint(Checkpoint.newBuilder()
                            .setSource(SOURCE)
                            .setConnectionId(connectionId)
                            .setScope(snapshot.getDriveId())
                            .setCursor(snapshot.getCursor()))
                    .build());
        } catch (RuntimeException e) {
            LOG.log(System.Logger.Level.WARNING, "sync-table snapshot failed: {0}", e.toString());
        }
    }

    @Override
    public void completeRun(String runId) {
        if (runId == null || runId.isBlank()) {
            return;
        }
        try {
            stub.reconcile(ReconcileRequest.newBuilder()
                    .setSource(SOURCE)
                    .setConnectionId(connectionId)
                    .setRunId(runId)
                    .build());
        } catch (RuntimeException e) {
            LOG.log(System.Logger.Level.WARNING, "sync-table reconcile failed: {0}", e.toString());
        }
    }

    static Asset toAsset(MicrosoftChange change) {
        return toAsset(DEFAULT_CONNECTION_ID, change);
    }

    static Asset toAsset(String connectionId, MicrosoftChange change) {
        MicrosoftEntity entity = change.getEntity();
        String kind = kind(entity);
        String connection = connectionId == null || connectionId.isBlank()
                ? DEFAULT_CONNECTION_ID : connectionId;
        Asset.Builder asset = Asset.newBuilder()
                .setAssetId(assetId(connection, entity))
                .setSource(SOURCE)
                .setConnectionId(connection)
                .setNativeId(entity.getEntityId())
                .setKind(kind)
                .setPhase(phase(change))
                .setStatus(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED)
                .setRunId(change.getCursor())
                .setCursor(change.getCursor());
        if (entity.hasIngestedAt()) {
            asset.setLastModifiedAt(entity.getIngestedAt());
        }
        if (entity.getEntityCase() == MicrosoftEntity.EntityCase.DRIVE_ITEM) {
            DriveItem item = entity.getDriveItem();
            asset.setTitle(item.getName())
                    .setSourceUri(item.getWebUrl())
                    .setContentBytes(item.getSize())
                    .setMediaType(item.getMimeType())
                    .setAttachment(!item.getFolder());
            if (!item.getParentId().isEmpty()) {
                asset.setParentAssetId(SOURCE + ":" + connection + ":drive_item:" + item.getParentId());
            }
        } else if (entity.getEntityCase() == MicrosoftEntity.EntityCase.DRIVE) {
            asset.setTitle(entity.getDrive().getName()).setSourceUri(entity.getDrive().getWebUrl());
        } else if (entity.getEntityCase() == MicrosoftEntity.EntityCase.SITE) {
            asset.setTitle(entity.getSite().getDisplayName()).setSourceUri(entity.getSite().getWebUrl());
        }
        return asset.build();
    }

    static String assetId(MicrosoftEntity entity) {
        return assetId(DEFAULT_CONNECTION_ID, entity);
    }

    static String assetId(String connectionId, MicrosoftEntity entity) {
        String connection = connectionId == null || connectionId.isBlank()
                ? DEFAULT_CONNECTION_ID : connectionId;
        return SOURCE + ":" + connection + ":" + kind(entity) + ":" + entity.getEntityId();
    }

    private static String kind(MicrosoftEntity entity) {
        return entity.getEntityCase().name().toLowerCase(Locale.ROOT);
    }

    private static AssetPhase phase(MicrosoftChange change) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            return AssetPhase.ASSET_PHASE_DELETE;
        }
        return change.getSource() == ChangeSource.CHANGE_SOURCE_CRAWL
                ? AssetPhase.ASSET_PHASE_INITIAL_CRAWL
                : AssetPhase.ASSET_PHASE_UPDATE;
    }

    /** Shuts down the underlying channel. */
    @Override
    public void close() {
        channel.shutdown();
        try {
            if (!channel.awaitTermination(5, TimeUnit.SECONDS)) {
                channel.shutdownNow();
            }
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            channel.shutdownNow();
        }
    }
}
