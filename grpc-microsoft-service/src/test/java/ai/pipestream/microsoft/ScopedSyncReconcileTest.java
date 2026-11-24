package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.sync.AssetStore;
import ai.pipestream.sync.SyncTableGrpcService;
import ai.pipestream.sync.v1.Asset;
import ai.pipestream.sync.v1.AssetSyncStatus;
import ai.pipestream.sync.v1.GetAssetRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import ai.pipestream.sync.v1.UpsertAssetRequest;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A sink may treat what a pass did not emit as gone only when the pass
 * enumerated everything the connection covers. A Sync whose request narrows
 * the scope (some drives, some sites, a folder), or that swept nothing, says
 * nothing about the rest, so it must not reconcile the ledger or replace the
 * mirror.
 */
class ScopedSyncReconcileTest {

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
            fake.stub("/drives/" + drive + "/root/delta", MicrosoftFixtures.deltaJson(
                    fake.baseUrl() + "/drives/" + drive + "/root/delta?token=t1",
                    MicrosoftFixtures.fileJson("file-in-" + drive, "notes.txt", drive)));
        }
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
    void aSyncOfOneDriveOutOfTwoLeavesTheOtherDrivesLedgerRowsAlone() throws Exception {
        String name = InProcessServerBuilder.generateName();
        syncServer = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(new SyncTableGrpcService(new AssetStore()))
                .build()
                .start();
        syncChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
        SyncTableServiceGrpc.SyncTableServiceBlockingStub ledger =
                SyncTableServiceGrpc.newBlockingStub(syncChannel);
        // Both drives were synced by an earlier run.
        for (String drive : List.of("drive-a", "drive-b")) {
            ledger.upsertAsset(UpsertAssetRequest.newBuilder()
                    .setAsset(Asset.newBuilder()
                            .setAssetId("microsoft:default:drive_item:file-in-" + drive)
                            .setSource("microsoft")
                            .setConnectionId("default")
                            .setKind("drive_item")
                            .setNativeId("file-in-" + drive)
                            .setRunId("earlier-run")
                            .setAttachment(true)
                            .setTitle("notes.txt"))
                    .build());
        }
        start(twoDrives().build(), new SyncTableMicrosoftChangeSink(syncChannel));

        sync(SyncRequest.newBuilder().addDriveIds("drive-a").build());

        assertThat(fake.requestsTo("/drives/drive-b/root/delta")).isEmpty();
        // The pass never looked at drive-b, so it cannot have found its file gone.
        assertThat(ledger.getAsset(GetAssetRequest.newBuilder()
                .setAssetId("microsoft:default:drive_item:file-in-drive-b").build())
                .getAsset().getStatus())
                .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
    }

    @Test
    void aSyncOfSomeOfTheConnectionsDrivesIsNotACompletePass() throws Exception {
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(twoDrives().build(), sink);

        sync(SyncRequest.newBuilder().addDriveIds("drive-a").build());

        assertThat(sink.fullSweeps()).containsExactly(false);
    }

    @Test
    void aSyncOfASiteTheRequestNamesIsNotACompletePass() throws Exception {
        fake.stub("/sites/site-1", MicrosoftFixtures.siteJson("site-1", "Team"));
        fake.stub("/sites/site-1/drives",
                MicrosoftFixtures.drivesJson(MicrosoftFixtures.driveJson("drive-a", "drive-a")));
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(twoDrives().build(), sink);

        sync(SyncRequest.newBuilder().addSiteIds("site-1").build());

        assertThat(fake.requestsTo("/drives/drive-b/root/delta")).isEmpty();
        assertThat(sink.fullSweeps()).containsExactly(false);
    }

    @Test
    void aSyncOfAFolderTheRequestNamesIsNotACompletePass() throws Exception {
        fake.stub("/drives/drive-a/root:/Shared:", "{\"id\":\"folder-1\"}");
        fake.stub("/drives/drive-a/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-a/root/delta?token=t1",
                MicrosoftFixtures.rootJson("drive-a"),
                MicrosoftFixtures.fileJson("file-in-drive-a", "notes.txt", "drive-a"),
                MicrosoftFixtures.folderJson("folder-1", "Shared", "drive-a")));
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(base().driveIds(List.of("drive-a")).build(), sink);

        sync(SyncRequest.newBuilder().setFolderPath("/Shared").build());

        // Everything outside /Shared was filtered out, not found missing.
        assertThat(sink.fullSweeps()).containsExactly(false);
    }

    @Test
    void aSyncThatSweptNothingIsNotACompletePass() throws Exception {
        fake.stub("/sites/site-1", MicrosoftFixtures.siteJson("site-1", "Team"));
        fake.stub("/sites/site-1/drives", MicrosoftFixtures.drivesJson());
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(base().siteId("site-1").build(), sink);

        sync(SyncRequest.getDefaultInstance());

        // No drive was enumerated, so nothing was shown to be gone.
        assertThat(sink.fullSweeps()).containsExactly(false);
    }

    @Test
    void aSyncOfTheWholeConnectionIsACompletePass() throws Exception {
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(twoDrives().build(), sink);

        sync(SyncRequest.getDefaultInstance());

        assertThat(fake.requestsTo("/drives/drive-b/root/delta")).hasSize(1);
        assertThat(sink.fullSweeps()).containsExactly(true);
        // It asked for no lists and no security, so it says nothing about them.
        assertThat(sink.enumeratedKinds()).containsExactly(Set.of(
                MicrosoftEntity.EntityCase.SITE, MicrosoftEntity.EntityCase.DRIVE,
                MicrosoftEntity.EntityCase.DRIVE_ITEM));
    }

    @Test
    void aRequestThatNamesExactlyTheConnectionsScopeIsACompletePass() throws Exception {
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        start(twoDrives().build(), sink);

        // The same drives in another order, and the configured folder spelled out.
        sync(SyncRequest.newBuilder().addDriveIds("drive-b").addDriveIds("drive-a")
                .setFolderPath("/").build());

        assertThat(sink.fullSweeps()).containsExactly(true);
    }

    private MicrosoftConnectorConfig.Builder base() {
        return MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl());
    }

    private MicrosoftConnectorConfig.Builder twoDrives() {
        return base().driveIds(List.of("drive-a", "drive-b"));
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

    private void sync(SyncRequest request) {
        MicrosoftServiceGrpc.newBlockingStub(microsoftChannel).sync(request)
                .forEachRemaining(ignored -> {
                });
    }
}
