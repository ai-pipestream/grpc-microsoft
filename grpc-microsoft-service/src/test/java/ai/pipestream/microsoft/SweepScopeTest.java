package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftEntity.EntityCase;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class SweepScopeTest {

    @Test
    void aDriveHoldsItsDriveItemsAndAListItsListItems() {
        SweepScope drive = SweepScope.drive("drive-1");
        SweepScope list = SweepScope.list("site-1", "list-1");

        assertThat(drive.kind()).isEqualTo(EntityCase.DRIVE_ITEM);
        assertThat(drive.id()).isEqualTo("drive-1");
        assertThat(list.kind()).isEqualTo(EntityCase.LIST_ITEM);
        assertThat(list.id()).isEqualTo("site-1/list-1");
        assertThat(drive.holds(EntityCase.DRIVE_ITEM, "drive-1")).isTrue();
        assertThat(drive.holds(EntityCase.DRIVE_ITEM, "drive-2")).isFalse();
        assertThat(drive.holds(EntityCase.LIST_ITEM, "drive-1")).isFalse();
        assertThat(drive.holds(EntityCase.DRIVE_ITEM, "")).isFalse();
    }

    @Test
    void theScopeOfAnEntityComesFromTheEntityAlone() {
        assertThat(SweepScope.idOf(MicrosoftEntity.newBuilder()
                .setDriveItem(DriveItem.newBuilder().setId("f").setDriveId("drive-1")).build()))
                .isEqualTo("drive-1");
        assertThat(SweepScope.idOf(MicrosoftEntity.newBuilder()
                .setListItem(ListItem.newBuilder().setId("1").setSiteId("site-1")
                        .setListId("list-1")).build()))
                .isEqualTo("site-1/list-1");
        // A list item that does not name its site belongs to no scope.
        assertThat(SweepScope.idOf(MicrosoftEntity.newBuilder()
                .setListItem(ListItem.newBuilder().setId("1").setListId("list-1")).build()))
                .isEmpty();
        assertThat(SweepScope.idOf(MicrosoftEntity.newBuilder()
                .setDrive(Drive.newBuilder().setId("drive-1")).build())).isEmpty();
        assertThat(SweepScope.idOf(MicrosoftEntity.getDefaultInstance())).isEmpty();
    }

    @Test
    void aScopeIsADriveOrAListWithAnId() {
        assertThatThrownBy(() -> new SweepScope(EntityCase.SITE, "site-1"))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("drive items or list items");
        assertThatThrownBy(() -> SweepScope.drive(" "))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("needs an id");
    }

    @Test
    void theLedgerRowOfAnItemCarriesItsScope() {
        assertThat(SyncTableMicrosoftChangeSink.toAsset(change(MicrosoftEntity.newBuilder()
                .setEntityId("f")
                .setDriveItem(DriveItem.newBuilder().setId("f").setDriveId("drive-1"))))
                .getScope()).isEqualTo("drive-1");
        assertThat(SyncTableMicrosoftChangeSink.toAsset(change(MicrosoftEntity.newBuilder()
                .setEntityId("list-1/1")
                .setListItem(ListItem.newBuilder().setId("1").setSiteId("site-1")
                        .setListId("list-1"))))
                .getScope()).isEqualTo("site-1/list-1");
        assertThat(SyncTableMicrosoftChangeSink.toAsset(change(MicrosoftEntity.newBuilder()
                .setEntityId("drive-1")
                .setDrive(Drive.newBuilder().setId("drive-1"))))
                .getScope()).isEmpty();
    }

    private static MicrosoftChange change(MicrosoftEntity.Builder entity) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setCursor("run-1")
                .setEntity(entity)
                .build();
    }
}
