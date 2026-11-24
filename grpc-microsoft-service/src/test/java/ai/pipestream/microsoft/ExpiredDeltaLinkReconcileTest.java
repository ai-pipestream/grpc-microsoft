package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.output.OutputFormat;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import ai.pipestream.sync.AssetStore;
import ai.pipestream.sync.SyncTableGrpcService;
import ai.pipestream.sync.v1.AssetSyncStatus;
import ai.pipestream.sync.v1.GetAssetRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import com.google.protobuf.Message;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * An incremental pass over two drives in which Graph rejects one drive's
 * delta link with {@code 410 Gone}. That drive is read again from scratch,
 * which sends no delete records: whatever was removed from it since the last
 * good pass is simply absent. The other drive stays on its delta link. The
 * pass must remove the first drive's absent items from the ledger and from
 * the output bundle, and must leave the second drive's unchanged items alone.
 */
class ExpiredDeltaLinkReconcileTest {

    private FakeGraphServer fake;
    private Server microsoftServer;
    private ManagedChannel microsoftChannel;
    private Server syncServer;
    private ManagedChannel syncChannel;

    @BeforeEach
    void startGraph() throws Exception {
        fake = FakeGraphServer.start();
        for (String drive : List.of("drive-a", "drive-b")) {
            fake.stub("/drives/" + drive, MicrosoftFixtures.driveJson(drive, drive));
        }
        // The first pass: both drives from scratch.
        fake.stub("/drives/drive-a/root/delta", MicrosoftFixtures.deltaJson(link("drive-a", "a1"),
                MicrosoftFixtures.fileJson("a-kept", "kept.txt", "drive-a"),
                MicrosoftFixtures.fileJson("a-gone", "gone.txt", "drive-a")));
        fake.stub("/drives/drive-b/root/delta", MicrosoftFixtures.deltaJson(link("drive-b", "b1"),
                MicrosoftFixtures.fileJson("b-unchanged", "same.txt", "drive-b")));
        // The second pass: drive-a's link has expired and a-gone is no longer
        // there; drive-b has no changes.
        fake.stub("/drives/drive-a/root/delta?token=a1", FakeGraphServer.error(410,
                "resyncChangesApplyDifferences", Map.of("Location", link("drive-a", "fresh"))));
        fake.stub("/drives/drive-a/root/delta?token=fresh", MicrosoftFixtures.deltaJson(
                link("drive-a", "a2"),
                MicrosoftFixtures.fileJson("a-kept", "kept.txt", "drive-a")));
        fake.stub("/drives/drive-b/root/delta?token=b1",
                MicrosoftFixtures.deltaJson(link("drive-b", "b2")));
    }

    @AfterEach
    void stop() {
        if (microsoftChannel != null) {
            microsoftChannel.shutdownNow();
        }
        if (microsoftServer != null) {
            microsoftServer.shutdownNow();
        }
        if (syncChannel != null) {
            syncChannel.shutdownNow();
        }
        if (syncServer != null) {
            syncServer.shutdownNow();
        }
        fake.close();
    }

    @Test
    void theLedgerLosesTheItemsGoneFromTheDriveThatWasReadAgain() throws Exception {
        startLedger(new SyncTableGrpcService(new AssetStore()));
        start(new SyncTableMicrosoftChangeSink(syncChannel));

        String cursor = sync(SyncRequest.getDefaultInstance());
        assertThat(status("a-gone")).isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        sync(SyncRequest.newBuilder().setSinceCursor(cursor).build());

        assertThat(fake.requestsTo("/drives/drive-a/root/delta")).extracting(r -> r.query())
                .contains("token=a1", "token=fresh");
        assertThat(status("a-gone")).as("the file removed from the drive that was read again")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        assertThat(status("a-kept")).isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        assertThat(status("b-unchanged")).as("the unchanged file of the drive still on delta")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
    }

    @Test
    void aLedgerRowThatNeverSaidItsDriveIsLeftToACompletePass() throws Exception {
        startLedger(new SyncTableGrpcService(new AssetStore()));
        start(new SyncTableMicrosoftChangeSink(syncChannel));
        // Written before rows carried a scope: nothing says which drive it
        // belongs to, so a pass that read drive-a alone cannot call it gone.
        SyncTableServiceGrpc.newBlockingStub(syncChannel).upsertAsset(
                ai.pipestream.sync.v1.UpsertAssetRequest.newBuilder()
                        .setAsset(ai.pipestream.sync.v1.Asset.newBuilder()
                                .setAssetId("microsoft:default:drive_item:legacy")
                                .setSource("microsoft")
                                .setConnectionId("default")
                                .setKind("drive_item")
                                .setNativeId("legacy")
                                .setRunId("earlier-run"))
                        .build());

        String cursor = sync(SyncRequest.newBuilder().addDriveIds("drive-a").build());
        assertThat(status("legacy")).as("after a pass of one drive out of two")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        // Every row the pass wrote says its drive.
        assertThat(SyncTableServiceGrpc.newBlockingStub(syncChannel)
                .getAsset(GetAssetRequest.newBuilder()
                        .setAssetId("microsoft:default:drive_item:a-kept").build())
                .getAsset().getScope()).isEqualTo("drive-a");

        sync(SyncRequest.getDefaultInstance());
        assertThat(status("legacy")).as("after a pass of the whole connection")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        assertThat(cursor).isNotEmpty();
    }

    @Test
    void aLedgerWithoutScopedReconcileFailsThePassAndDeletesNothing() throws Exception {
        // A sync-table built before ReconcileScope existed.
        startLedger(io.grpc.ServerInterceptors.intercept(
                new SyncTableGrpcService(new AssetStore()), new io.grpc.ServerInterceptor() {
                    @Override
                    public <Q, R> io.grpc.ServerCall.Listener<Q> interceptCall(
                            io.grpc.ServerCall<Q, R> call, io.grpc.Metadata headers,
                            io.grpc.ServerCallHandler<Q, R> next) {
                        if ("ReconcileScope".equals(call.getMethodDescriptor().getBareMethodName())) {
                            call.close(io.grpc.Status.UNIMPLEMENTED.withDescription(
                                    "Method not found: " + call.getMethodDescriptor()
                                            .getFullMethodName()), new io.grpc.Metadata());
                            return new io.grpc.ServerCall.Listener<>() {
                            };
                        }
                        return next.startCall(call, headers);
                    }
                }));
        start(new SyncTableMicrosoftChangeSink(syncChannel));
        String cursor = sync(SyncRequest.getDefaultInstance());

        List<SyncResponse> received = new ArrayList<>();
        java.util.Iterator<SyncResponse> second = MicrosoftServiceGrpc
                .newBlockingStub(microsoftChannel)
                .sync(SyncRequest.newBuilder().setSinceCursor(cursor).build());
        org.assertj.core.api.Assertions.assertThatThrownBy(() -> second.forEachRemaining(received::add))
                .isInstanceOfSatisfying(io.grpc.StatusRuntimeException.class, e ->
                        assertThat(e.getStatus().getDescription())
                                .contains("sync-table").contains("ReconcileScope")
                                .contains("drive-a").contains("UNIMPLEMENTED"));

        assertThat(received).noneMatch(
                r -> r.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR);
        // Nothing was reconciled more widely in its place.
        assertThat(status("a-gone")).isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        assertThat(status("b-unchanged")).isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
    }

    @Test
    void theLedgerLosesTheItemsGoneFromAListThatWasReadAgain() throws Exception {
        String lists = "/sites/site-1/lists";
        fake.stub("/sites/site-1", MicrosoftFixtures.siteJson("site-1", "Team"));
        fake.stub("/sites/site-1/drives",
                MicrosoftFixtures.drivesJson(MicrosoftFixtures.driveJson("drive-b", "drive-b")));
        fake.stub(lists, "{\"value\":[{\"id\":\"L1\",\"name\":\"Tasks\","
                + "\"list\":{\"template\":\"genericList\"}}]}");
        String listLink = fake.baseUrl() + lists + "/L1/items/delta?token=";
        fake.stub(lists + "/L1/items/delta", MicrosoftFixtures.deltaJson(listLink + "l1",
                "{\"id\":\"1\",\"fields\":{\"Title\":\"kept\"}}",
                "{\"id\":\"2\",\"fields\":{\"Title\":\"gone\"}}"));
        // The second pass: the list's link has expired and item 2 is gone.
        fake.stub(lists + "/L1/items/delta?token=l1", FakeGraphServer.error(410,
                "resyncChangesApplyDifferences", Map.of("Location", listLink + "fresh")));
        fake.stub(lists + "/L1/items/delta?token=fresh", MicrosoftFixtures.deltaJson(listLink + "l2",
                "{\"id\":\"1\",\"fields\":{\"Title\":\"kept\"}}"));
        startLedger(new SyncTableGrpcService(new AssetStore()));
        InMemoryMicrosoftChangeSink seen = new InMemoryMicrosoftChangeSink();
        start(base().siteId("site-1").build(), new CompositeMicrosoftChangeSink(List.of(
                new SyncTableMicrosoftChangeSink(syncChannel), seen)));

        String cursor = sync(SyncRequest.newBuilder().setIncludeLists(true).build());
        sync(SyncRequest.newBuilder().setIncludeLists(true).setSinceCursor(cursor).build());

        // The first pass was complete; the second read only the list in full.
        assertThat(seen.enumeratedScopes()).containsExactly(Set.of(),
                Set.of(SweepScope.list("site-1", "L1")));
        assertThat(seen.fullSweeps()).containsExactly(true, false);
        assertThat(rowStatus("list_item", "L1/2")).as("the item gone from the list")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        assertThat(rowStatus("list_item", "L1/1"))
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        assertThat(rowStatus("list", "L1")).isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
        assertThat(status("b-unchanged")).as("the unchanged file of the drive still on delta")
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_SYNCED);
    }

    @Test
    void thePassNamesTheDriveItReadAgainAndNoKind() throws Exception {
        InMemoryMicrosoftChangeSink seen = new InMemoryMicrosoftChangeSink();
        start(seen);

        String cursor = sync(SyncRequest.getDefaultInstance());
        sync(SyncRequest.newBuilder().setSinceCursor(cursor).build());

        // A complete pass needs no scopes: its kinds cover every drive.
        assertThat(seen.fullSweeps()).containsExactly(true, false);
        assertThat(seen.enumeratedScopes()).containsExactly(Set.of(),
                Set.of(SweepScope.drive("drive-a")));
    }

    @Test
    void theOutputBundleLosesTheItemsGoneFromTheDriveThatWasReadAgain() throws Exception {
        List<String> bundles = new ArrayList<>();
        start(new OutputMicrosoftChangeSink(nullStore(), List.of(recorder(bundles)), "out"));

        String cursor = sync(SyncRequest.getDefaultInstance());
        sync(SyncRequest.newBuilder().setSinceCursor(cursor).build());

        assertThat(bundles).containsExactly(
                "out=[a-gone, a-kept, b-unchanged, drive-a, drive-b]",
                "out=[a-kept, b-unchanged, drive-a, drive-b]");
    }

    private AssetSyncStatus status(String itemId) {
        return rowStatus("drive_item", itemId);
    }

    private AssetSyncStatus rowStatus(String kind, String nativeId) {
        return SyncTableServiceGrpc.newBlockingStub(syncChannel)
                .getAsset(GetAssetRequest.newBuilder()
                        .setAssetId("microsoft:default:" + kind + ":" + nativeId).build())
                .getAsset().getStatus();
    }

    private void startLedger(io.grpc.BindableService ledger) throws Exception {
        startLedger(ledger.bindService());
    }

    private void startLedger(io.grpc.ServerServiceDefinition ledger) throws Exception {
        String name = InProcessServerBuilder.generateName();
        syncServer = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(ledger)
                .build()
                .start();
        syncChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
    }

    private String link(String drive, String token) {
        return fake.baseUrl() + "/drives/" + drive + "/root/delta?token=" + token;
    }

    private MicrosoftConnectorConfig.Builder base() {
        return MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl());
    }

    private void start(MicrosoftChangeSink downstream) throws Exception {
        start(base().driveIds(List.of("drive-a", "drive-b")).build(), downstream);
    }

    private void start(MicrosoftConnectorConfig config, MicrosoftChangeSink downstream)
            throws Exception {
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String name = InProcessServerBuilder.generateName();
        microsoftServer = InProcessServerBuilder.forName(name)
                // The service blocks on outbound flow control, as in production:
                // a direct executor would run it on the client thread and stall.
                .executor(java.util.concurrent.Executors.newVirtualThreadPerTaskExecutor())
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES, downstream))
                .build()
                .start();
        microsoftChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
    }

    /** Runs one Sync and returns its resume cursor. */
    private String sync(SyncRequest request) {
        String[] cursor = {""};
        MicrosoftServiceGrpc.newBlockingStub(microsoftChannel).sync(request)
                .forEachRemaining(response -> {
                    if (response.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                        cursor[0] = response.getResumeCursor();
                    }
                });
        assertThat(cursor[0]).isNotEmpty();
        return cursor[0];
    }

    private static OutputFormat recorder(List<String> bundles) {
        return new OutputFormat() {
            @Override
            public String id() {
                return "recorder";
            }

            @Override
            public boolean supports(Message record) {
                return false;
            }

            @Override
            public void complete(OutputStore store, List<Message> records, String prefix) {
                bundles.add(prefix + "=" + records.stream()
                        .map(r -> ((MicrosoftChange) r).getEntity().getEntityId()).sorted().toList());
            }
        };
    }

    private static OutputStore nullStore() {
        return new OutputStore() {
            @Override
            public String id() {
                return "null";
            }

            @Override
            public boolean available(Map<String, String> env) {
                return true;
            }

            @Override
            public void open(Map<String, String> env) {
            }

            @Override
            public void put(OutputObject object) {
            }
        };
    }
}
