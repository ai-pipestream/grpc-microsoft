package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import microsoft.graph.connectors.contracts.grpc.ContentItem;
import microsoft.graph.connectors.contracts.grpc.CrawlItem;
import microsoft.graph.connectors.contracts.grpc.DeletedItem;
import microsoft.graph.connectors.contracts.grpc.GenericType;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlItem;
import microsoft.graph.connectors.contracts.grpc.ScdPrincipal;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyValueMap;

import java.util.Optional;

/**
 * {@link MicrosoftChange} / {@link DriveItem} onto GCA crawl items.
 * Folders and files are both {@code ContentItem}s so one Sync stream
 * finishes the crawl without relying on GCA to re-enter for LinkItems.
 *
 * <p>A file's bytes become the item's content only when they are text the
 * connector can decode and send (see {@link FileContent}). When a crawl
 * asked for content and a file's content is not on its item, the mapping
 * says why in {@link Mapped#omittedContent()}.
 */
public final class CrawlItemMapper {

    /**
     * Largest file content sent on an item, in bytes of UTF-8. Stays under
     * GCA's 4 MiB crawl-item ceiling, and under the 4 MiB a Sync response
     * may be, so the crawl also asks Sync not to inline anything larger.
     */
    static final int CONTENT_CAP_BYTES = 3 * 1024 * 1024;

    private CrawlItemMapper() {
    }

    /**
     * A mapped crawl item and what became of its file content.
     *
     * @param item the GCA item
     * @param omittedContent present when the crawl asked for content and
     *        this file's content is not on the item
     * @param <T> {@link CrawlItem} or {@link IncrementalCrawlItem}
     */
    public record Mapped<T>(T item, Optional<OmittedContent> omittedContent) {
    }

    /**
     * Maps a non-delete {@link MicrosoftChange} to a full-crawl {@link CrawlItem}.
     * Deletes are skipped; GCA full crawls do not carry deleted items.
     *
     * @param change a Sync change event
     * @param contentRequested whether the Sync was asked to inline file bytes
     * @return the crawl item, or empty when the change is a delete or not a
     *         drive item
     */
    public static Optional<Mapped<CrawlItem>> toCrawlItem(MicrosoftChange change,
            boolean contentRequested) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            return Optional.empty();
        }
        return driveItem(change).map(item -> toCrawlItem(item, contentRequested));
    }

    /**
     * Maps a {@link MicrosoftChange} to an incremental crawl item, including
     * deletes.
     *
     * @param change a Sync change event
     * @param contentRequested whether the Sync was asked to inline file bytes
     * @return the incremental item, or empty when the entity is missing or not a
     *         drive item
     */
    public static Optional<Mapped<IncrementalCrawlItem>> toIncrementalCrawlItem(
            MicrosoftChange change, boolean contentRequested) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            String itemId = change.getEntity().getEntityId();
            if (itemId.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new Mapped<>(IncrementalCrawlItem.newBuilder()
                    .setItemType(IncrementalCrawlItem.ItemType.DeletedItem)
                    .setItemId(itemId)
                    .setDeletedItem(DeletedItem.getDefaultInstance())
                    .build(), Optional.empty()));
        }
        return driveItem(change).map(item -> {
            FileContent.Outcome content = FileContent.of(item, contentRequested);
            return new Mapped<>(IncrementalCrawlItem.newBuilder()
                    .setItemType(IncrementalCrawlItem.ItemType.ContentItem)
                    .setItemId(itemId(item))
                    .setContentItem(contentItem(item, content))
                    .build(), content.omitted());
        });
    }

    /**
     * Maps a {@link DriveItem} to a GCA {@link CrawlItem} content item.
     *
     * @param item the drive item
     * @param contentRequested whether the Sync was asked to inline file bytes
     * @return the crawl item and what became of its content
     */
    public static Mapped<CrawlItem> toCrawlItem(DriveItem item, boolean contentRequested) {
        FileContent.Outcome content = FileContent.of(item, contentRequested);
        return new Mapped<>(CrawlItem.newBuilder()
                .setItemType(CrawlItem.ItemType.ContentItem)
                .setItemId(itemId(item))
                .setContentItem(contentItem(item, content))
                .build(), content.omitted());
    }

    private static Optional<DriveItem> driveItem(MicrosoftChange change) {
        if (!change.hasEntity()) {
            return Optional.empty();
        }
        MicrosoftEntity entity = change.getEntity();
        if (entity.getEntityCase() != MicrosoftEntity.EntityCase.DRIVE_ITEM) {
            return Optional.empty();
        }
        return Optional.of(entity.getDriveItem());
    }

    static String itemId(DriveItem item) {
        if (item.getDriveId().isEmpty()) {
            return item.getId();
        }
        return item.getDriveId() + "/" + item.getId();
    }

    private static ContentItem contentItem(DriveItem item, FileContent.Outcome content) {
        ContentItem.Builder builder = ContentItem.newBuilder()
                .setPropertyValues(properties(item));
        content.content().ifPresent(builder::setContent);
        return builder.build();
    }

    private static SourcePropertyValueMap properties(DriveItem item) {
        SourcePropertyValueMap.Builder values = SourcePropertyValueMap.newBuilder();
        putString(values, DataSourceSchemas.TITLE, item.getName());
        putString(values, DataSourceSchemas.FILE_NAME, item.getName());
        putString(values, DataSourceSchemas.FILE_EXTENSION, extension(item.getName()));
        putString(values, DataSourceSchemas.WEB_URL, item.getWebUrl());
        if (item.hasCreatedAt()) {
            values.putValues(DataSourceSchemas.CREATED,
                    GenericType.newBuilder().setDateTimeValue(item.getCreatedAt()).build());
        }
        if (item.hasLastModifiedAt()) {
            values.putValues(DataSourceSchemas.LAST_MODIFIED,
                    GenericType.newBuilder().setDateTimeValue(item.getLastModifiedAt()).build());
        }
        principal(item.getCreatedBy()).ifPresent(p -> values.putValues(DataSourceSchemas.CREATED_BY,
                GenericType.newBuilder().setPrincipalValue(p).build()));
        principal(item.getLastModifiedBy()).ifPresent(p ->
                values.putValues(DataSourceSchemas.LAST_MODIFIED_BY,
                        GenericType.newBuilder().setPrincipalValue(p).build()));
        putString(values, DataSourceSchemas.MIME_TYPE, item.getMimeType());
        values.putValues(DataSourceSchemas.SIZE,
                GenericType.newBuilder().setIntValue(item.getSize()).build());
        putString(values, DataSourceSchemas.PARENT_ID, item.getParentId());
        putString(values, DataSourceSchemas.DRIVE_ID, item.getDriveId());
        values.putValues(DataSourceSchemas.FOLDER,
                GenericType.newBuilder().setBoolValue(item.getFolder()).build());
        putString(values, DataSourceSchemas.ITEM_TYPE, item.getFolder() ? "folder" : "file");
        return values.build();
    }

    private static Optional<ScdPrincipal> principal(IdentitySet set) {
        Identity identity = set.getUser();
        if (identity.getId().isEmpty() && identity.getDisplayName().isEmpty()) {
            return Optional.empty();
        }
        return Optional.of(ScdPrincipal.newBuilder()
                .setExternalId(identity.getId())
                .setExternalName(identity.getDisplayName())
                .setEntraDisplayName(identity.getDisplayName())
                .setEntraId(identity.getId())
                .build());
    }

    private static void putString(SourcePropertyValueMap.Builder values, String name, String value) {
        if (value == null || value.isEmpty()) {
            return;
        }
        values.putValues(name, GenericType.newBuilder().setStringValue(value).build());
    }

    private static String extension(String name) {
        if (name == null) {
            return "";
        }
        int dot = name.lastIndexOf('.');
        if (dot <= 0 || dot == name.length() - 1) {
            return "";
        }
        return name.substring(dot + 1);
    }
}
