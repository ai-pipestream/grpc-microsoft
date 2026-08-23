package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import microsoft.graph.connectors.contracts.grpc.Content;
import microsoft.graph.connectors.contracts.grpc.ContentItem;
import microsoft.graph.connectors.contracts.grpc.CrawlItem;
import microsoft.graph.connectors.contracts.grpc.DeletedItem;
import microsoft.graph.connectors.contracts.grpc.GenericType;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlItem;
import microsoft.graph.connectors.contracts.grpc.ScdPrincipal;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyValueMap;

import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.Optional;

/**
 * {@link MicrosoftChange} / {@link DriveItem} onto GCA crawl items.
 * Folders and files are both {@code ContentItem}s so one Sync stream
 * finishes the crawl without relying on GCA to re-enter for LinkItems.
 */
public final class CrawlItemMapper {

    /** Stay under GCA's 4 MiB crawl-item ceiling. */
    static final int CONTENT_CAP_BYTES = 3 * 1024 * 1024;

    private CrawlItemMapper() {
    }

    public static Optional<CrawlItem> toCrawlItem(MicrosoftChange change) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            return Optional.empty();
        }
        return driveItem(change).map(CrawlItemMapper::toCrawlItem);
    }

    public static Optional<IncrementalCrawlItem> toIncrementalCrawlItem(MicrosoftChange change) {
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            String itemId = change.getEntity().getEntityId();
            if (itemId.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(IncrementalCrawlItem.newBuilder()
                    .setItemType(IncrementalCrawlItem.ItemType.DeletedItem)
                    .setItemId(itemId)
                    .setDeletedItem(DeletedItem.getDefaultInstance())
                    .build());
        }
        return driveItem(change).map(CrawlItemMapper::toIncrementalContent);
    }

    public static CrawlItem toCrawlItem(DriveItem item) {
        return CrawlItem.newBuilder()
                .setItemType(CrawlItem.ItemType.ContentItem)
                .setItemId(itemId(item))
                .setContentItem(contentItem(item))
                .build();
    }

    private static IncrementalCrawlItem toIncrementalContent(DriveItem item) {
        return IncrementalCrawlItem.newBuilder()
                .setItemType(IncrementalCrawlItem.ItemType.ContentItem)
                .setItemId(itemId(item))
                .setContentItem(contentItem(item))
                .build();
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

    private static ContentItem contentItem(DriveItem item) {
        ContentItem.Builder builder = ContentItem.newBuilder()
                .setPropertyValues(properties(item));
        content(item).ifPresent(builder::setContent);
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

    private static Optional<Content> content(DriveItem item) {
        if (!item.hasContent() || item.getContent().isEmpty()) {
            return Optional.empty();
        }
        byte[] bytes = item.getContent().toByteArray();
        if (bytes.length > CONTENT_CAP_BYTES) {
            return Optional.empty();
        }
        String mime = item.getMimeType().toLowerCase(Locale.ROOT);
        if (!(mime.startsWith("text/") || mime.contains("json") || mime.contains("xml")
                || mime.contains("html"))) {
            return Optional.empty();
        }
        Content.ContentType type = mime.contains("html")
                ? Content.ContentType.Html : Content.ContentType.Text;
        return Optional.of(Content.newBuilder()
                .setContentType(type)
                .setContentValue(new String(bytes, StandardCharsets.UTF_8))
                .build());
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
