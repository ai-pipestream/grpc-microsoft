package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftEntity.EntityCase;
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

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A complete pass shows that an entity is gone only for the kinds of entity
 * it enumerated. Lists and list items are read only with
 * {@code include_lists}, site groups, admins and ACLs only with
 * {@code include_security}. A full Sync without those flags never looked at
 * them, so it must leave the rows an earlier Sync wrote for them alone.
 */
class KindScopedReconcileTest {

    private static final List<String> LIST_KINDS = List.of("list", "list_item");
    private static final List<String> SECURITY_KINDS =
            List.of("site_group", "site_admins", "site_acl");

    private FakeGraphServer fake;
    private Server microsoftServer;
    private ManagedChannel microsoftChannel;
    private Server syncServer;
    private ManagedChannel syncChannel;
    private SyncTableServiceGrpc.SyncTableServiceBlockingStub ledger;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        fake.stub("/sites/site-1", MicrosoftFixtures.siteJson("site-1", "Team"));
        fake.stub("/sites/site-1/drives",
                MicrosoftFixtures.drivesJson(MicrosoftFixtures.driveJson("drive-a", "Documents")));
        fake.stub("/drives/drive-a/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-a/root/delta?token=t1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-a")));
        // The site has no non-library list any more.
        fake.stub("/sites/site-1/lists", "{\"value\":[]}");

        String name = InProcessServerBuilder.generateName();
        syncServer = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(new SyncTableGrpcService(new AssetStore()))
                .build()
                .start();
        syncChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
        ledger = SyncTableServiceGrpc.newBlockingStub(syncChannel);
        // An earlier Sync with include_lists and include_security wrote these.
        for (String kind : LIST_KINDS) {
            earlier(kind, kind + "-1");
        }
        for (String kind : SECURITY_KINDS) {
            earlier(kind, kind + "-1");
        }
        // And a file that has since been deleted.
        earlier("drive_item", "file-gone");

        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl())
                .siteId("site-1")
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String microsoftName = InProcessServerBuilder.generateName();
        microsoftServer = InProcessServerBuilder.forName(microsoftName)
                // The service blocks on outbound flow control, as in production:
                // a direct executor would run it on the client thread and stall.
                .executor(java.util.concurrent.Executors.newVirtualThreadPerTaskExecutor())
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES,
                        new SyncTableMicrosoftChangeSink(syncChannel)))
                .build()
                .start();
        microsoftChannel = InProcessChannelBuilder.forName(microsoftName).directExecutor().build();
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
    void aFullSyncWithoutListsOrSecurityLeavesTheirRowsAlone() {
        sync(SyncRequest.getDefaultInstance());

        // The pass read no list and no security data.
        assertThat(fake.requestsTo("/sites/site-1/lists")).isEmpty();
        for (String kind : LIST_KINDS) {
            assertThat(status(kind, kind + "-1")).as(kind)
                    .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        }
        for (String kind : SECURITY_KINDS) {
            assertThat(status(kind, kind + "-1")).as(kind)
                    .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        }
    }

    @Test
    void aFullSyncStillFindsADeletedFileGone() {
        sync(SyncRequest.getDefaultInstance());

        assertThat(status("drive_item", "file-gone"))
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        assertThat(status("drive_item", "file-1"))
                .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
    }

    @Test
    void aFullSyncWithListsFindsDeletedListsGoneAndLeavesSecurityRowsAlone() {
        sync(SyncRequest.newBuilder().setIncludeLists(true).build());

        assertThat(fake.requestsTo("/sites/site-1/lists")).isNotEmpty();
        // The pass enumerated the site's lists and found none.
        for (String kind : LIST_KINDS) {
            assertThat(status(kind, kind + "-1")).as(kind)
                    .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        }
        // It still read no security data.
        for (String kind : SECURITY_KINDS) {
            assertThat(status(kind, kind + "-1")).as(kind)
                    .isNotEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
        }
    }

    @Test
    void theKindsACompletePassEnumeratesFollowWhatItRead() {
        // No site in scope (a drive allowlist): lists and site security are
        // never read; the ACLs of the drives' libraries are all there is.
        assertThat(MicrosoftCrawler.enumeratedKinds(false, false, false))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM);
        assertThat(MicrosoftCrawler.enumeratedKinds(false, true, false))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM);
        assertThat(MicrosoftCrawler.enumeratedKinds(false, false, true))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM,
                        EntityCase.SITE_ACL);
        // A site in scope.
        assertThat(MicrosoftCrawler.enumeratedKinds(true, false, false))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM);
        assertThat(MicrosoftCrawler.enumeratedKinds(true, true, false))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM,
                        EntityCase.LIST, EntityCase.LIST_ITEM);
        // Security without lists leaves the ACLs of the site's lists unread,
        // so ACLs as a kind were not enumerated.
        assertThat(MicrosoftCrawler.enumeratedKinds(true, false, true))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM,
                        EntityCase.SITE_GROUP, EntityCase.SITE_ADMINS);
        assertThat(MicrosoftCrawler.enumeratedKinds(true, true, true))
                .containsExactlyInAnyOrder(EntityCase.SITE, EntityCase.DRIVE, EntityCase.DRIVE_ITEM,
                        EntityCase.LIST, EntityCase.LIST_ITEM, EntityCase.SITE_GROUP,
                        EntityCase.SITE_ADMINS, EntityCase.SITE_ACL);
        // Users and permissions are never part of a crawl.
        assertThat(MicrosoftCrawler.enumeratedKinds(true, true, true))
                .doesNotContain(EntityCase.USER, EntityCase.PERMISSION);
    }

    private void earlier(String kind, String nativeId) {
        ledger.upsertAsset(UpsertAssetRequest.newBuilder()
                .setAsset(Asset.newBuilder()
                        .setAssetId(assetId(kind, nativeId))
                        .setSource("microsoft")
                        .setConnectionId("default")
                        .setKind(kind)
                        .setNativeId(nativeId)
                        .setRunId("earlier-run")
                        .setTitle(nativeId))
                .build());
    }

    private AssetSyncStatus status(String kind, String nativeId) {
        return ledger.getAsset(GetAssetRequest.newBuilder().setAssetId(assetId(kind, nativeId))
                .build()).getAsset().getStatus();
    }

    private static String assetId(String kind, String nativeId) {
        // The ledger names a row's kind after its MicrosoftEntity case.
        assertThat(MicrosoftEntity.EntityCase.valueOf(kind.toUpperCase(java.util.Locale.ROOT)))
                .isNotNull();
        return "microsoft:default:" + kind + ":" + nativeId;
    }

    private void sync(SyncRequest request) {
        MicrosoftServiceGrpc.newBlockingStub(microsoftChannel).sync(request)
                .forEachRemaining(ignored -> {
                });
    }
}
