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
        assertThat(folderAsset.getAssetId()).isEqualTo("microsoft:drive_item:folder-1");
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
        assertThat(fileAsset.getParentAssetId()).isEqualTo("microsoft:drive_item:folder-1");
        assertThat(fileAsset.getPhase())
                .isEqualTo(ai.pipestream.sync.v1.AssetPhase.ASSET_PHASE_UPDATE);
        assertThat(fileAsset.getContentBytes()).isEqualTo(12);
    }
}
