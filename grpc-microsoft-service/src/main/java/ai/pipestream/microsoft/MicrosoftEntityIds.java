package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.PermissionTarget;

/**
 * {@code MicrosoftEntity.entity_id} values for entities whose Graph id is not
 * unique on its own. List item ids are small integers per list, so a list item
 * is keyed by {@code list_id/item_id}; a permission id repeats on every item
 * that inherits the grant, so a permission is keyed by its target plus its id.
 */
public final class MicrosoftEntityIds {

    private MicrosoftEntityIds() {
    }

    /**
     * The entity id of a list item.
     *
     * @param item the list item (with {@code list_id} set)
     * @return {@code list_id/item_id}
     */
    public static String listItem(ListItem item) {
        return item.getListId() + "/" + item.getId();
    }

    /**
     * The entity id of a permission.
     *
     * @param permission the permission (with {@code target} set)
     * @return {@code target-key/permission-id}
     */
    public static String permission(Permission permission) {
        return target(permission.getTarget()) + "/" + permission.getId();
    }

    /**
     * A stable key for a permission target.
     *
     * @param target the target
     * @return for example {@code drive_item/{drive}/{item}} or {@code site/{site}}
     */
    public static String target(PermissionTarget target) {
        return switch (target.getTargetCase()) {
            case DRIVE_ITEM -> "drive_item/" + target.getDriveItem().getDriveId() + "/"
                    + target.getDriveItem().getItemId();
            case LIST_ITEM -> "list_item/" + target.getListItem().getSiteId() + "/"
                    + target.getListItem().getListId() + "/" + target.getListItem().getItemId();
            case LIST -> "list/" + target.getList().getSiteId() + "/" + target.getList().getListId();
            case SITE_ID -> "site/" + target.getSiteId();
            case TARGET_NOT_SET -> "unknown";
        };
    }
}
