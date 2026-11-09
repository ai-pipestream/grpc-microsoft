package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.output.ObjectKeys;

/**
 * Store keys organized by Microsoft Graph hierarchy. Segments are
 * filesystem-safe; result has no leading slash.
 */
public final class MicrosoftObjectKeys {

    private MicrosoftObjectKeys() {
    }

    /**
     * Hierarchy key for a change, or {@code records/{hash}} when the
     * change carries no entity.
     *
     * @param change a crawler change
     * @return store key without extension
     */
    public static String of(MicrosoftChange change) {
        if (change.hasEntity()) {
            return of(change.getEntity());
        }
        return ObjectKeys.join("records", Integer.toHexString(System.identityHashCode(change)));
    }

    /**
     * Graph hierarchy: {@code sites/{id}}, {@code sites/{site}/drives/{id}},
     * {@code drives/{driveId}/items/{id}}, {@code users/{id}}.
     *
     * @param entity wrapped entity
     * @return key without extension
     */
    public static String of(MicrosoftEntity entity) {
        return switch (entity.getEntityCase()) {
            case SITE -> {
                Site site = entity.getSite();
                yield ObjectKeys.join("sites", id(site.getId(), entity));
            }
            case DRIVE -> {
                Drive drive = entity.getDrive();
                yield drive.getSiteId().isBlank()
                        ? ObjectKeys.join("drives", id(drive.getId(), entity))
                        : ObjectKeys.join("sites", ObjectKeys.segment(drive.getSiteId()), "drives",
                                id(drive.getId(), entity));
            }
            case DRIVE_ITEM -> {
                DriveItem item = entity.getDriveItem();
                yield ObjectKeys.join("drives", ObjectKeys.segment(item.getDriveId()), "items",
                        id(item.getId(), entity));
            }
            case USER -> {
                GraphUser user = entity.getUser();
                yield ObjectKeys.join("users", id(user.getId(), entity));
            }
            default -> ObjectKeys.join("entities", id(entity.getEntityId(), entity));
        };
    }


    private static String id(String nativeId, MicrosoftEntity entity) {
        if (nativeId != null && !nativeId.isBlank()) {
            return ObjectKeys.segment(nativeId);
        }
        return ObjectKeys.segment(entity.getEntityId());
    }
}
