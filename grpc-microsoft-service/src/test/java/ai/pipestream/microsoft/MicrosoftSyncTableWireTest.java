package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.sync.AssetStore;
import ai.pipestream.sync.SyncTableGrpcService;
import ai.pipestream.sync.v1.Asset;
import ai.pipestream.sync.v1.AssetSyncStatus;
import ai.pipestream.sync.v1.GetAssetRequest;
import ai.pipestream.sync.v1.ListAssetsRequest;
import ai.pipestream.sync.v1.SyncTableServiceGrpc;
import ai.pipestream.sync.v1.UpsertAssetRequest;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftSyncTableWireTest {

    private FakeGraphServer fake;
    private Server microsoftServer;
    private Server syncServer;
    private ManagedChannel microsoftChannel;
    private ManagedChannel syncChannel;

    @AfterEach
    void stop() {
        if (microsoftChannel != null) {
            microsoftChannel.shutdownNow();
        }
        if (syncChannel != null) {
            syncChannel.shutdownNow();
        }
        if (microsoftServer != null) {
            microsoftServer.shutdownNow();
        }
        if (syncServer != null) {
            syncServer.shutdownNow();
        }
        if (fake != null) {
            fake.close();
        }
    }

    @Test
    void syncWritesFilesAndReconcilesStaleRows() throws Exception {
        fake = FakeGraphServer.start();
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/children", MicrosoftFixtures.childrenJson(
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));

        String syncName = InProcessServerBuilder.generateName();
        String msName = InProcessServerBuilder.generateName();
        syncServer = InProcessServerBuilder.forName(syncName)
                .directExecutor()
                .addService(new SyncTableGrpcService(new AssetStore()))
                .build()
                .start();
        syncChannel = InProcessChannelBuilder.forName(syncName).directExecutor().build();
        SyncTableServiceGrpc.SyncTableServiceBlockingStub ledger =
                SyncTableServiceGrpc.newBlockingStub(syncChannel);
        ledger.upsertAsset(UpsertAssetRequest.newBuilder()
                .setAsset(Asset.newBuilder()
                        .setAssetId("microsoft:drive_item:gone")
                        .setSource("microsoft")
                        .setKind("drive_item")
                        .setNativeId("gone")
                        .setRunId("stale-run")
                        .setAttachment(true)
                        .setTitle("removed.txt"))
                .build());

        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        microsoftServer = InProcessServerBuilder.forName(msName)
                .directExecutor()
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES,
                        new SyncTableMicrosoftChangeSink(syncChannel)))
                .build()
                .start();
        microsoftChannel = InProcessChannelBuilder.forName(msName).directExecutor().build();
        MicrosoftServiceGrpc.newBlockingStub(microsoftChannel)
                .sync(SyncRequest.getDefaultInstance())
                .forEachRemaining(ignored -> {
                });

        Asset file = ledger.getAsset(GetAssetRequest.newBuilder()
                .setAssetId("microsoft:drive_item:file-1").build()).getAsset();
        assertThat(file.getAttachment()).isTrue();
        assertThat(file.getTitle()).isEqualTo("notes.txt");

        List<String> attachments = new ArrayList<>();
        ledger.listAssets(ListAssetsRequest.newBuilder()
                .setSource("microsoft")
                .setAttachmentsOnly(true)
                .build()).forEachRemaining(r -> attachments.add(r.getAsset().getAssetId()));
        assertThat(attachments).contains("microsoft:drive_item:file-1");

        assertThat(ledger.getAsset(GetAssetRequest.newBuilder()
                .setAssetId("microsoft:drive_item:gone").build()).getAsset().getStatus())
                .isEqualTo(AssetSyncStatus.ASSET_SYNC_STATUS_DELETED);
    }
}
