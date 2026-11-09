package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.Site;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftObjectKeysTest {

    @Test
    void microsoftHierarchyUsesSiteDriveItem() {
        assertThat(MicrosoftObjectKeys.of(MicrosoftEntity.newBuilder()
                .setEntityId("site-1")
                .setSite(Site.newBuilder().setId("site-1"))
                .build())).isEqualTo("sites/site-1");
        assertThat(MicrosoftObjectKeys.of(MicrosoftEntity.newBuilder()
                .setEntityId("drive-1")
                .setDrive(Drive.newBuilder().setId("drive-1").setSiteId("site-1"))
                .build())).isEqualTo("sites/site-1/drives/drive-1");
        assertThat(MicrosoftObjectKeys.of(MicrosoftEntity.newBuilder()
                .setEntityId("file-1")
                .setDriveItem(DriveItem.newBuilder().setId("file-1").setDriveId("drive-1"))
                .build())).isEqualTo("drives/drive-1/items/file-1");
    }

    @Test
    void ofUsesChangeEntity() {
        assertThat(MicrosoftObjectKeys.of(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("file-1")
                        .setDriveItem(DriveItem.newBuilder().setId("file-1").setDriveId("d")))
                .build())).isEqualTo("drives/d/items/file-1");
    }
}
