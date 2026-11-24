package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.sync.AssetStore;
import ai.pipestream.sync.SyncTableGrpcService;
import ai.pipestream.sync.v1.Asset;
import ai.pipestream.sync.v1.AssetSyncStatus;
import ai.pipestream.sync.v1.GetAssetRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import ai.pipestream.sync.v1.UpsertAssetRequest;
import io.grpc.ManagedChannel;
import io.grpc.Metadata;
import io.grpc.Server;
import io.grpc.ServerCall;
import io.grpc.ServerCallHandler;
import io.grpc.ServerInterceptor;
import io.grpc.ServerInterceptors;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * A sync-table write that fails is the sink's failure, not a log line: the
 * Sync that made it ends with an error and hands out no resume cursor, so the
 * delta link does not move past a change the ledger never recorded, and the
 * ledger is not reconciled against a run whose rows did not all land.
 */
class SyncTableLedgerFailureTest {

    /** The real ledger behind a switch that fails chosen RPCs and counts every call. */
    private static final class Ledger implements ServerInterceptor {
        final Map<String, Status> failing = new ConcurrentHashMap<>();
        final Map<String, Integer> calls = new ConcurrentHashMap<>();

        @Override
        public <Q, R> ServerCall.Listener<Q> interceptCall(ServerCall<Q, R> call,
                Metadata headers, ServerCallHandler<Q, R> next) {
            String method = call.getMethodDescriptor().getBareMethodName();
            calls.merge(method, 1, Integer::sum);
            Status failure = failing.get(method);
            if (failure != null) {
                call.close(failure, new Metadata());
                return new ServerCall.Listener<>() {
                };
            }
            return next.startCall(call, headers);
        }

        void fail(String method) {
            failing.put(method, Status.UNAVAILABLE.withDescription("ledger restarting"));
        }

        int calls(String method) {
            return calls.getOrDefault(method, 0);
        }
    }

    private final Ledger ledger = new Ledger();
    private Server syncServer;
    private ManagedChannel syncChannel;
    private FakeGraphServer fake;
    private Server microsoftServer;
    private ManagedChannel microsoftChannel;

    @BeforeEach
    void startLedger() throws Exception {
        String name = InProcessServerBuilder.generateName();
        syncServer = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(ServerInterceptors.intercept(
                        new SyncTableGrpcService(new AssetStore()), ledger))
                .build()
                .start();
        syncChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
    }

    @AfterEach
    void stop() {
        if (microsoftChannel != null) {
            microsoftChannel.shutdownNow();
        }
        if (microsoftServer != null) {
            microsoftServer.shutdownNow();
        }
        if (fake != null) {
            fake.close();
        }
        syncChannel.shutdownNow();
        syncServer.shutdownNow();
    }

    @Test
    void aFailedUpsertIsThrownAndNamesTheAsset() {
        ledger.fail("UpsertAsset");
        SyncTableMicrosoftChangeSink sink = new SyncTableMicrosoftChangeSink(syncChannel);

        assertThatThrownBy(() -> sink.emit(fileChange("file-1", ChangeOperation.CHANGE_OPERATION_UPSERT)))
                .isInstanceOf(SyncTableMicrosoftChangeSink.WriteFailedException.class)
                .hasMessageContaining("sync-table")
                .hasMessageContaining("UpsertAsset")
                .hasMessageContaining("microsoft:default:drive_item:file-1")
                .hasMessageContaining("ledger restarting");
    }

    @Test
    void aFailedDeleteIsThrownAndNamesTheAsset() {
        ledger.fail("DeleteAsset");
        SyncTableMicrosoftChangeSink sink = new SyncTableMicrosoftChangeSink(syncChannel);

        assertThatThrownBy(() -> sink.emit(fileChange("file-1", ChangeOperation.CHANGE_OPERATION_DELETE)))
                .isInstanceOf(SyncTableMicrosoftChangeSink.WriteFailedException.class)
                .hasMessageContaining("sync-table")
                .hasMessageContaining("DeleteAsset")
                .hasMessageContaining("microsoft:default:drive_item:file-1");
    }

    @Test
    void aFailedCheckpointAndAFailedReconcileAreThrown() {
        ledger.fail("PutCheckpoint");
        ledger.fail("Reconcile");
        SyncTableMicrosoftChangeSink sink = new SyncTableMicrosoftChangeSink(syncChannel);

        assertThatThrownBy(() -> sink.snapshot(MicrosoftSnapshot.newBuilder()
                .setDriveId("drive-1").setCursor("run-1").build()))
                .isInstanceOf(SyncTableMicrosoftChangeSink.WriteFailedException.class)
                .hasMessageContaining("sync-table")
                .hasMessageContaining("PutCheckpoint")
                .hasMessageContaining("drive-1");
        assertThatThrownBy(() -> sink.completeRun("run-1"))
                .isInstanceOf(SyncTableMicrosoftChangeSink.WriteFailedException.class)
                .hasMessageContaining("sync-table")
                .hasMessageContaining("Reconcile")
                .hasMessageContaining("run-1");
    }

    @Test
    void anIncrementalSyncWhoseLedgerWriteFailsHandsOutNoResumeCursor() throws Exception {
        fake = FakeGraphServer.start();
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta?token=t1", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=t2",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));
        startMicrosoft();
        ledger.fail("UpsertAsset");
        String since = SyncCursors.encode(SyncCursor.newBuilder()
                .putDriveDeltaLinks("drive-1", fake.baseUrl() + "/drives/drive-1/root/delta?token=t1")
                .build());

        List<SyncResponse> received = new ArrayList<>();
        Iterator<SyncResponse> sync = MicrosoftServiceGrpc.newBlockingStub(microsoftChannel)
                .sync(SyncRequest.newBuilder().setSinceCursor(since).build());
        assertThatThrownBy(() -> sync.forEachRemaining(received::add))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> {
                    assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.UNAVAILABLE);
                    assertThat(e.getStatus().getDescription())
                            .contains("sync-table").contains("UpsertAsset");
                });

        // The delta link must not advance past a change the ledger never
        // recorded: Graph reports it once, so the caller has to ask again
        // from the cursor it already holds.
        assertThat(received).noneMatch(
                r -> r.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR);
    }

    @Test
    void aFullSyncWhoseLedgerWritesFailEndsWithAnErrorAndDeletesNothing() throws Exception {
        // The file exists in the drive and in the ledger from an earlier run.
        SyncTableServiceGrpc.SyncTableServiceBlockingStub rows =
                SyncTableServiceGrpc.newBlockingStub(syncChannel);
        rows.upsertAsset(UpsertAssetRequest.newBuilder()
                .setAsset(Asset.newBuilder()
                        .setAssetId("microsoft:default:drive_item:file-1")
                        .setSource("microsoft")
                        .setConnectionId("default")
                        .setKind("drive_item")
                        .setNativeId("file-1")
                        .setRunId("earlier-run")
                        .setAttachment(true)
                        .setTitle("notes.txt"))
                .build());
        AssetSyncStatus before = status(rows, "file-1");
        fake = FakeGraphServer.start();
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=t1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));
        startMicrosoft();
        ledger.fail("UpsertAsset");

        List<SyncResponse> received = new ArrayList<>();
        Iterator<SyncResponse> sync = MicrosoftServiceGrpc.newBlockingStub(microsoftChannel)
                .sync(SyncRequest.getDefaultInstance());
        assertThatThrownBy(() -> sync.forEachRemaining(received::add))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> {
                    assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.UNAVAILABLE);
                    assertThat(e.getStatus().getDescription())
                            .contains("sync-table").contains("UpsertAsset");
                });

        assertThat(received).noneMatch(
                r -> r.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR);
        // The ledger is not reconciled against the run: the file whose
        // upsert failed still exists and must not be marked deleted.
        assertThat(ledger.calls("Reconcile")).isZero();
        assertThat(status(rows, "file-1"))
                .isEqualTo(before)
                .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
    }

    @Test
    void aFailedReconcileFailsTheSync() throws Exception {
        fake = FakeGraphServer.start();
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=t1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));
        startMicrosoft();
        ledger.fail("Reconcile");

        List<SyncResponse> received = new ArrayList<>();
        Iterator<SyncResponse> sync = MicrosoftServiceGrpc.newBlockingStub(microsoftChannel)
                .sync(SyncRequest.getDefaultInstance());
        assertThatThrownBy(() -> sync.forEachRemaining(received::add))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e ->
                        assertThat(e.getStatus().getDescription())
                                .contains("sync-table").contains("Reconcile"));
        assertThat(received).noneMatch(
                r -> r.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR);
    }

    @Test
    void aLedgerRefusalIsNotPassedOffAsGraphsAnswer() throws Exception {
        fake = FakeGraphServer.start();
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=t1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));
        startMicrosoft();
        ledger.failing.put("UpsertAsset",
                Status.PERMISSION_DENIED.withDescription("ledger says no"));

        Iterator<SyncResponse> sync = MicrosoftServiceGrpc.newBlockingStub(microsoftChannel)
                .sync(SyncRequest.getDefaultInstance());
        // PERMISSION_DENIED from Sync means Graph refused the caller; the
        // ledger's refusal is the service's own trouble.
        assertThatThrownBy(() -> sync.forEachRemaining(ignored -> {
        })).isInstanceOfSatisfying(StatusRuntimeException.class, e -> {
            assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INTERNAL);
            assertThat(e.getStatus().getDescription())
                    .contains("sync-table").contains("PERMISSION_DENIED").contains("ledger says no");
        });
    }

    private static AssetSyncStatus status(SyncTableServiceGrpc.SyncTableServiceBlockingStub rows,
            String nativeId) {
        return rows.getAsset(GetAssetRequest.newBuilder()
                .setAssetId("microsoft:default:drive_item:" + nativeId).build()).getAsset().getStatus();
    }

    private void startMicrosoft() throws Exception {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String name = InProcessServerBuilder.generateName();
        microsoftServer = InProcessServerBuilder.forName(name)
                // The service blocks on outbound flow control, as in production:
                // a direct executor would run it on the client thread and stall.
                .executor(java.util.concurrent.Executors.newVirtualThreadPerTaskExecutor())
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES,
                        new SyncTableMicrosoftChangeSink(syncChannel)))
                .build()
                .start();
        microsoftChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
    }

    private static MicrosoftChange fileChange(String id, ChangeOperation operation) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c-" + id)
                .setOperation(operation)
                .setSource(ChangeSource.CHANGE_SOURCE_DELTA)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(id)
                        .setDriveItem(DriveItem.newBuilder()
                                .setId(id)
                                .setName("notes.txt")
                                .setDriveId("drive-1")))
                .build();
    }
}
