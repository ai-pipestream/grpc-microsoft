package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.MicrosoftEntity;

import java.util.Objects;

/**
 * A part of a connection that one sweep reads on its own: the items of one
 * drive, or the items of one list. A pass names the scopes it enumerated in
 * full when it did not enumerate the whole connection: typically one drive
 * whose delta link Graph rejected with {@code 410 Gone}, read again from
 * scratch while the other drives stayed incremental. Within such a scope,
 * and only there, a sink may treat an entity it holds but the pass did not
 * emit as gone.
 *
 * @param kind the kind of entity the scope holds:
 *        {@link MicrosoftEntity.EntityCase#DRIVE_ITEM} for a drive,
 *        {@link MicrosoftEntity.EntityCase#LIST_ITEM} for a list
 * @param id the drive id, or {@code site_id/list_id} for a list. The same
 *        value as the scope of the sweep's checkpoint in the sync table
 */
public record SweepScope(MicrosoftEntity.EntityCase kind, String id) {

    /**
     * Checks the scope.
     *
     * @param kind drive items or list items
     * @param id the drive id or list key
     */
    public SweepScope {
        Objects.requireNonNull(kind, "kind");
        Objects.requireNonNull(id, "id");
        if (kind != MicrosoftEntity.EntityCase.DRIVE_ITEM
                && kind != MicrosoftEntity.EntityCase.LIST_ITEM) {
            throw new IllegalArgumentException("a sweep reads drive items or list items, not " + kind);
        }
        if (id.isBlank()) {
            throw new IllegalArgumentException("a sweep scope needs an id");
        }
    }

    /**
     * The items of one drive.
     *
     * @param driveId the drive id
     * @return the scope
     */
    public static SweepScope drive(String driveId) {
        return new SweepScope(MicrosoftEntity.EntityCase.DRIVE_ITEM, driveId);
    }

    /**
     * The items of one list.
     *
     * @param siteId the site id
     * @param listId the list id
     * @return the scope
     */
    public static SweepScope list(String siteId, String listId) {
        return new SweepScope(MicrosoftEntity.EntityCase.LIST_ITEM,
                SyncCursors.listKey(siteId, listId));
    }

    /**
     * The id of the scope an entity belongs to, worked out from the entity
     * alone so that every sink agrees on it.
     *
     * @param entity a Graph entity
     * @return the drive id of a drive item, {@code site_id/list_id} of a list
     *         item; empty for every other kind and for an item that does not
     *         name its drive or its site and list
     */
    public static String idOf(MicrosoftEntity entity) {
        return switch (entity.getEntityCase()) {
            case DRIVE_ITEM -> entity.getDriveItem().getDriveId();
            case LIST_ITEM -> {
                ListItem item = entity.getListItem();
                yield item.getSiteId().isEmpty() || item.getListId().isEmpty() ? ""
                        : SyncCursors.listKey(item.getSiteId(), item.getListId());
            }
            default -> "";
        };
    }

    /**
     * Whether an entity of {@code entityKind} whose scope id is
     * {@code scopeId} lies in this scope.
     *
     * @param entityKind the kind of the entity
     * @param scopeId the entity's {@link #idOf scope id}
     * @return true when it does
     */
    public boolean holds(MicrosoftEntity.EntityCase entityKind, String scopeId) {
        return kind == entityKind && id.equals(scopeId);
    }
}
