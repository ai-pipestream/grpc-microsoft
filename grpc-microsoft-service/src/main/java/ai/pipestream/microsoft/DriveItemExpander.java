package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemInclude;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.PermissionTarget;
import com.fasterxml.jackson.databind.JsonNode;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Applies a request's {@link DriveItemInclude} to drive items: the backing list
 * item ({@code /listItem?$expand=fields}: id, content type, typed columns), the
 * item's permissions ({@code /permissions}, all pages) and its retention label
 * ({@code /retentionLabel}); the publication facet is a {@code $select} on the
 * item read itself. Personal-OneDrive files have no list item and items without
 * a label answer 404; both leave the field empty instead of failing.
 */
final class DriveItemExpander {

    /**
     * Resolved expansion flags.
     *
     * @param listColumns fetch the list item and its columns
     * @param permissions fetch the permissions
     * @param retentionLabel fetch the retention label
     * @param publication select the publication facet
     */
    record Flags(boolean listColumns, boolean permissions, boolean retentionLabel,
            boolean publication) {

        /**
         * The flags of a request.
         *
         * @param present whether the request set its {@code include} message
         * @param include the request's {@code include}
         * @param defaultListColumns what an unset {@code include} means for list
         *        columns (the rpc's first-revision behaviour)
         * @return the flags
         */
        static Flags of(boolean present, DriveItemInclude include, boolean defaultListColumns) {
            if (!present) {
                return new Flags(defaultListColumns, false, false, false);
            }
            return new Flags(include.getListColumns(), include.getPermissions(),
                    include.getRetentionLabel(), include.getPublication());
        }
    }

    private final GraphFiles files;
    private final GraphPager pager;
    private final MicrosoftMapper mapper;
    private final Flags flags;

    DriveItemExpander(GraphFiles files, MicrosoftMapper mapper, Flags flags) {
        this.files = Objects.requireNonNull(files, "files");
        this.pager = new GraphPager(files.client());
        this.mapper = Objects.requireNonNull(mapper, "mapper");
        this.flags = Objects.requireNonNull(flags, "flags");
    }

    /** The {@code $select} for item reads under these flags. */
    String select() {
        return GraphPaths.driveItemSelect(flags.publication());
    }

    /** Candidate URLs for reading one item route: full select first, then Graph's default. */
    List<String> candidates(String route) {
        return List.of(route + "?$select=" + select(), route);
    }

    /** Adds the requested expansions to a (non-deleted) item. */
    DriveItem expand(DriveItem item) throws IOException, InterruptedException {
        if (item.hasDeleted()) {
            return item;
        }
        String driveId = item.getDriveId();
        String itemId = item.getId();
        if (flags.listColumns() && !item.getIsRoot()) {
            item = mapper.withListItem(item, orNull(() -> files.listItemFields(driveId, itemId)));
        }
        if (flags.permissions()) {
            PermissionTarget target = PermissionTarget.newBuilder()
                    .setDriveItem(DriveItemRef.newBuilder().setDriveId(driveId).setItemId(item.getId()))
                    .build();
            DriveItem.Builder b = item.toBuilder().clearPermissions();
            pager.forEach(List.of(GraphPaths.item(driveId, item.getId()) + "/permissions"), Map.of(),
                    node -> {
                        b.addPermissions(mapper.toPermission(node, target));
                        return true;
                    });
            item = b.build();
        }
        if (flags.retentionLabel() && !item.getFolder()) {
            String id = item.getId();
            JsonNode label = orNull(() -> files.client().get(
                    GraphPaths.item(driveId, id) + "/retentionLabel", Map.of()));
            if (label != null) {
                item = item.toBuilder().setRetentionLabel(mapper.toRetentionLabel(label)).build();
            }
        }
        return item;
    }

    /**
     * Downloads at most {@code max} bytes of an item's content.
     *
     * @param files the files API
     * @param driveId the drive
     * @param itemId the item
     * @param max the limit; 0 = none
     * @return the bytes, or {@code null} when the body is larger than {@code max}
     */
    static byte[] inline(GraphFiles files, String driveId, String itemId, long max)
            throws IOException, InterruptedException {
        try (GraphClient.Content content = files.client().openContent(
                GraphPaths.item(driveId, itemId) + "/content", "")) {
            InputStream in = content.body();
            if (max == 0) {
                return in.readAllBytes();
            }
            byte[] bytes = in.readNBytes((int) Math.min(Integer.MAX_VALUE - 8L, max + 1));
            return bytes.length > max ? null : bytes;
        }
    }

    @FunctionalInterface
    private interface GraphCall {
        JsonNode call() throws IOException, InterruptedException;
    }

    private static JsonNode orNull(GraphCall call) throws IOException, InterruptedException {
        try {
            return call.call();
        } catch (GraphClient.GraphApiException e) {
            if (e.status() == 404) {
                return null;
            }
            throw e;
        }
    }
}
