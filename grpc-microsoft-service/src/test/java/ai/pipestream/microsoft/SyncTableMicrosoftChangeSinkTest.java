package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.sync.v1.Asset;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class SyncTableMicrosoftChangeSinkTest {

    @Test
    void mapsFolderAndFile() {
        MicrosoftChange folder = MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setSource(ChangeSource.CHANGE_SOURCE_CRAWL)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("folder-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                        .setDriveItem(DriveItem.newBuilder()
                                .setId("folder-1")
                                .setName("Docs")
                                .setDriveId("drive-1")
                                .setFolder(true)
                                .setWebUrl("https://contoso.sharepoint.com/Docs")))
                .build();
        Asset folderAsset = SyncTableMicrosoftChangeSink.toAsset(folder);
        assertThat(folderAsset.getAssetId()).isEqualTo("microsoft:default:drive_item:folder-1");
        assertThat(folderAsset.getConnectionId()).isEqualTo("default");
        assertThat(folderAsset.getAttachment()).isFalse();
        assertThat(folderAsset.getPhase())
                .isEqualTo(ai.pipestream.sync.v1.AssetPhase.ASSET_PHASE_INITIAL_CRAWL);

        MicrosoftChange file = MicrosoftChange.newBuilder()
                .setChangeId("c2")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setSource(ChangeSource.CHANGE_SOURCE_DELTA)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                        .setDriveItem(DriveItem.newBuilder()
                                .setId("file-1")
                                .setName("notes.txt")
                                .setDriveId("drive-1")
                                .setParentId("folder-1")
                                .setFolder(false)
                                .setSize(12)
                                .setMimeType("text/plain")
                                .setWebUrl("https://contoso.sharepoint.com/notes.txt")))
                .build();
        Asset fileAsset = SyncTableMicrosoftChangeSink.toAsset(file);
        assertThat(fileAsset.getAttachment()).isTrue();
        assertThat(fileAsset.getParentAssetId()).isEqualTo("microsoft:default:drive_item:folder-1");
        assertThat(fileAsset.getPhase())
                .isEqualTo(ai.pipestream.sync.v1.AssetPhase.ASSET_PHASE_UPDATE);
        assertThat(fileAsset.getContentBytes()).isEqualTo(12);
    }

    @Test
    void mapsSiteDriveAndDelete() {
        Asset site = SyncTableMicrosoftChangeSink.toAsset(MicrosoftChange.newBuilder()
                .setChangeId("c3")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setSource(ChangeSource.CHANGE_SOURCE_CRAWL)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("site-1")
                        .setSite(ai.pipestream.microsoft.v1.Site.newBuilder()
                                .setId("site-1")
                                .setDisplayName("Docs")
                                .setWebUrl("https://contoso.sharepoint.com/sites/Docs")))
                .build());
        assertThat(site.getAssetId()).isEqualTo("microsoft:default:site:site-1");
        assertThat(site.getAttachment()).isFalse();

        Asset drive = SyncTableMicrosoftChangeSink.toAsset(MicrosoftChange.newBuilder()
                .setChangeId("c4")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setSource(ChangeSource.CHANGE_SOURCE_CRAWL)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("drive-1")
                        .setDrive(ai.pipestream.microsoft.v1.Drive.newBuilder()
                                .setId("drive-1")
                                .setName("Documents")
                                .setWebUrl("https://contoso.sharepoint.com/Documents")))
                .build());
        assertThat(drive.getKind()).isEqualTo("drive");
        assertThat(drive.getTitle()).isEqualTo("Documents");

        Asset deleted = SyncTableMicrosoftChangeSink.toAsset(MicrosoftChange.newBuilder()
                .setChangeId("c5")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setSource(ChangeSource.CHANGE_SOURCE_DELTA)
                .setCursor("run-1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setDriveItem(DriveItem.newBuilder().setId("file-1").setName("gone.txt")))
                .build());
        assertThat(deleted.getPhase())
                .isEqualTo(ai.pipestream.sync.v1.AssetPhase.ASSET_PHASE_DELETE);
    }
}
