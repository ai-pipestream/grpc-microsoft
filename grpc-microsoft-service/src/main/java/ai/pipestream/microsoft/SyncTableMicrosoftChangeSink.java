package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.ListItem;
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
import ai.pipestream.sync.v1.ReconcileScopeRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import ai.pipestream.sync.v1.UpsertAssetRequest;
import io.grpc.ManagedChannel;
import io.grpc.ManagedChannelBuilder;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;

import java.util.Locale;
import java.util.Set;
import java.util.TreeSet;
import java.util.Objects;
import java.util.concurrent.TimeUnit;

/**
 * Writes Microsoft Graph changes into {@code SyncTableService}. A write the
 * ledger refuses or cannot take is thrown as {@link WriteFailedException};
 * nothing is logged and skipped.
 */
public final class SyncTableMicrosoftChangeSink implements MicrosoftChangeSink, AutoCloseable {

    /** Asset {@code source} written for every Microsoft Graph row. */
    public static final String SOURCE = "microsoft";
    /** Catalog id used when the caller does not name a connection. */
    public static final String DEFAULT_CONNECTION_ID = "default";
    /** Environment variable for the {@code SyncTableService} target. */
    public static final String ENV_TARGET = "SYNC_TABLE_TARGET";
    /** Environment variable; {@code false} disables plaintext on the channel. */
    public static final String ENV_PLAINTEXT = "SYNC_TABLE_PLAINTEXT";

    private final ManagedChannel channel;
    private final SyncTableServiceGrpc.SyncTableServiceBlockingStub stub;
    private final String connectionId;
    /** False for a {@link #forConnection} view, which shares (and never closes) the channel. */
    private final boolean ownsChannel;

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
        this(channel, connectionId, true);
    }

    private SyncTableMicrosoftChangeSink(ManagedChannel channel, String connectionId,
            boolean ownsChannel) {
        this.channel = Objects.requireNonNull(channel, "channel");
        this.stub = SyncTableServiceGrpc.newBlockingStub(channel);
        this.connectionId = connectionId == null || connectionId.isBlank()
                ? DEFAULT_CONNECTION_ID : connectionId;
        this.ownsChannel = ownsChannel;
    }

    /**
     * A view of this sink that stamps {@code connectionId}; it shares the channel
     * and its {@link #close()} leaves the channel open.
     *
     * @param connectionId catalog connection
     * @return the bound view
     */
    @Override
    public MicrosoftChangeSink forConnection(String connectionId) {
        if (connectionId == null || connectionId.isBlank() || connectionId.equals(this.connectionId)) {
            return this;
        }
        return new SyncTableMicrosoftChangeSink(channel, connectionId, false);
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

    /**
     * A sync-table write failed. The ledger does not hold what the pass
     * emitted, so the pass that made the call must fail: its resume cursor
     * would move the delta link past a change the ledger never recorded, and
     * a reconcile would mark the rows whose upserts were lost as deleted.
     */
    public static final class WriteFailedException extends RuntimeException {

        private static final long serialVersionUID = 1L;

        /** The ledger's own status code for the failed call. */
        private final Status.Code ledgerCode;

        private WriteFailedException(String rpc, String subject, StatusRuntimeException cause) {
            super("sync-table " + rpc + " failed for " + subject + ": "
                    + cause.getStatus().getCode()
                    + (cause.getStatus().getDescription() == null ? ""
                            : ": " + cause.getStatus().getDescription()), cause);
            this.ledgerCode = cause.getStatus().getCode();
        }

        /**
         * The status code the sync-table answered with.
         *
         * @return the ledger's status code, never {@code null}
         */
        public Status.Code ledgerStatus() {
            return ledgerCode;
        }
    }

    /**
     * Writes the change to the ledger.
     *
     * @param change the change to record
     * @throws WriteFailedException if the sync-table refuses or cannot be reached
     */
    @Override
    public void emit(MicrosoftChange change) {
        String assetId = assetId(connectionId, change.getEntity());
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            try {
                stub.deleteAsset(DeleteAssetRequest.newBuilder()
                        .setAssetId(assetId)
                        .setRunId(change.getCursor())
                        .setCursor(change.getCursor())
                        .build());
            } catch (StatusRuntimeException e) {
                throw new WriteFailedException("DeleteAsset", assetId, e);
            }
            return;
        }
        try {
            stub.upsertAsset(UpsertAssetRequest.newBuilder()
                    .setAsset(toAsset(connectionId, change)).build());
        } catch (StatusRuntimeException e) {
            throw new WriteFailedException("UpsertAsset", assetId, e);
        }
    }

    /**
     * Records the sweep's checkpoint.
     *
     * @param snapshot the marker of a completed drive or list sweep
     * @throws WriteFailedException if the sync-table refuses or cannot be reached
     */
    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        String scope = scope(snapshot);
        try {
            stub.putCheckpoint(PutCheckpointRequest.newBuilder()
                    .setCheckpoint(Checkpoint.newBuilder()
                            .setSource(SOURCE)
                            .setConnectionId(connectionId)
                            .setScope(scope)
                            .setCursor(snapshot.getCursor()))
                    .build());
        } catch (StatusRuntimeException e) {
            throw new WriteFailedException("PutCheckpoint",
                    "scope " + scope + " of connection " + connectionId, e);
        }
    }

    /**
     * Checkpoint scope of a snapshot: the drive id for a drive sweep,
     * {@code site_id/list_id} for a list sweep.
     */
    static String scope(MicrosoftSnapshot snapshot) {
        if (!snapshot.getDriveId().isEmpty()) {
            return snapshot.getDriveId();
        }
        if (!snapshot.getListId().isEmpty()) {
            return SyncCursors.listKey(snapshot.getSiteId(), snapshot.getListId());
        }
        return snapshot.getSiteId();
    }

    /**
     * Reconciles by run id, one kind at a time: rows of this connection and
     * of a kind in {@code enumerated} that the run did not write are marked
     * deleted. Rows of every other kind are left as they are, because the
     * pass did not read that kind. An incremental or narrowed pass names no
     * kind: its unchanged rows keep their older run id.
     *
     * <p>Such a pass may still have read single drives or lists again from
     * scratch. For each of those {@code scopes} the rows stamped with that
     * scope that the run did not write are marked deleted, through
     * {@code ReconcileScope}. A ledger too old to have that rpc answers
     * {@code UNIMPLEMENTED}, which fails the pass: nothing is reconciled
     * more widely in its place. Rows written before rows carried a scope
     * are outside every scope and are only reconciled by a complete pass.</p>
     *
     * @param runId identifier of the completed run
     * @param enumerated the kinds of entity the pass enumerated in full
     * @param scopes the drives and lists the pass enumerated in full on
     *        their own
     * @throws WriteFailedException if the sync-table refuses or cannot be reached
     */
    @Override
    public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated,
            Set<SweepScope> scopes) {
        if (runId == null || runId.isBlank()) {
            return;
        }
        for (MicrosoftEntity.EntityCase kind : new TreeSet<>(enumerated)) {
            if (kind == MicrosoftEntity.EntityCase.ENTITY_NOT_SET) {
                continue;
            }
            try {
                stub.reconcile(ReconcileRequest.newBuilder()
                        .setSource(SOURCE)
                        .setConnectionId(connectionId)
                        .setRunId(runId)
                        .setKind(kind(kind))
                        .build());
            } catch (StatusRuntimeException e) {
                throw new WriteFailedException("Reconcile", "kind " + kind(kind) + " of run "
                        + runId + " of connection " + connectionId, e);
            }
        }
        for (SweepScope scope : scopes) {
            if (enumerated.contains(scope.kind())) {
                // The whole kind was reconciled above.
                continue;
            }
            try {
                stub.reconcileScope(ReconcileScopeRequest.newBuilder()
                        .setSource(SOURCE)
                        .setConnectionId(connectionId)
                        .setRunId(runId)
                        .setKind(kind(scope.kind()))
                        .setScope(scope.id())
                        .build());
            } catch (StatusRuntimeException e) {
                throw new WriteFailedException("ReconcileScope", "scope " + scope.id() + " ("
                        + kind(scope.kind()) + ") of run " + runId + " of connection "
                        + connectionId, e);
            }
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
                .setScope(SweepScope.idOf(entity))
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
        } else if (entity.getEntityCase() == MicrosoftEntity.EntityCase.LIST) {
            asset.setTitle(entity.getList().getDisplayName()).setSourceUri(entity.getList().getWebUrl());
        } else if (entity.getEntityCase() == MicrosoftEntity.EntityCase.LIST_ITEM) {
            ListItem item = entity.getListItem();
            asset.setTitle(item.getName()).setSourceUri(item.getWebUrl());
            if (!item.getListId().isEmpty()) {
                asset.setParentAssetId(SOURCE + ":" + connection + ":list:" + item.getListId());
            }
        } else if (entity.getEntityCase() == MicrosoftEntity.EntityCase.PERMISSION) {
            asset.setTitle(String.join(",", entity.getPermission().getRolesList()));
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
        return kind(entity.getEntityCase());
    }

    /** The ledger's name for a kind of entity. */
    static String kind(MicrosoftEntity.EntityCase kind) {
        return kind.name().toLowerCase(Locale.ROOT);
    }

    private static AssetPhase phase(MicrosoftChange change) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            return AssetPhase.ASSET_PHASE_DELETE;
        }
        return change.getSource() == ChangeSource.CHANGE_SOURCE_CRAWL
                ? AssetPhase.ASSET_PHASE_INITIAL_CRAWL
                : AssetPhase.ASSET_PHASE_UPDATE;
    }

    /** Shuts down the underlying channel (a {@link #forConnection} view leaves it open). */
    @Override
    public void close() {
        if (!ownsChannel) {
            return;
        }
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
