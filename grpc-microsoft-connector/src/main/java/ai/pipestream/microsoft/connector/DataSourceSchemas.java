package ai.pipestream.microsoft.connector;

import microsoft.graph.connectors.contracts.grpc.DataSourceSchema;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyDefinition;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyDefinition.SearchAnnotations;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyDefinition.SearchPropertyLabel;
import microsoft.graph.connectors.contracts.grpc.SourcePropertyDefinition.SourcePropertyType;

/**
 * Drive-item schema GCA stores on the connection. Property names are the
 * keys {@link CrawlItemMapper} writes into {@code SourcePropertyValueMap}.
 */
public final class DataSourceSchemas {

    public static final String TITLE = "Title";
    public static final String FILE_NAME = "FileName";
    public static final String FILE_EXTENSION = "FileExtension";
    public static final String WEB_URL = "WebUrl";
    public static final String CREATED = "Created";
    public static final String LAST_MODIFIED = "LastModified";
    public static final String CREATED_BY = "CreatedBy";
    public static final String LAST_MODIFIED_BY = "LastModifiedBy";
    public static final String MIME_TYPE = "MimeType";
    public static final String SIZE = "Size";
    public static final String PARENT_ID = "ParentId";
    public static final String DRIVE_ID = "DriveId";
    public static final String FOLDER = "Folder";
    public static final String ITEM_TYPE = "ItemType";

    private static final int SEARCHABLE_QUERYABLE_RETRIEVABLE =
            SearchAnnotations.IsSearchable_VALUE
                    | SearchAnnotations.IsQueryable_VALUE
                    | SearchAnnotations.IsRetrievable_VALUE;
    private static final int QUERYABLE_RETRIEVABLE =
            SearchAnnotations.IsQueryable_VALUE | SearchAnnotations.IsRetrievable_VALUE;

    private DataSourceSchemas() {
    }

    public static DataSourceSchema driveItem() {
        return DataSourceSchema.newBuilder()
                .addPropertyList(property(TITLE, SourcePropertyType.String,
                        SEARCHABLE_QUERYABLE_RETRIEVABLE, SearchPropertyLabel.Title,
                        "Display name of the drive item"))
                .addPropertyList(property(FILE_NAME, SourcePropertyType.String,
                        SEARCHABLE_QUERYABLE_RETRIEVABLE, SearchPropertyLabel.FileName,
                        "File or folder name"))
                .addPropertyList(property(FILE_EXTENSION, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.FileExtension,
                        "File extension without the leading dot"))
                .addPropertyList(property(WEB_URL, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.Url,
                        "Browser URL of the item"))
                .addPropertyList(property(CREATED, SourcePropertyType.DateTime,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.CreatedDateTime,
                        "When the item was created"))
                .addPropertyList(property(LAST_MODIFIED, SourcePropertyType.DateTime,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.LastModifiedDateTime,
                        "When the item was last modified"))
                .addPropertyList(property(CREATED_BY, SourcePropertyType.Principal,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.CreatedBy,
                        "User who created the item"))
                .addPropertyList(property(LAST_MODIFIED_BY, SourcePropertyType.Principal,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.LastModifiedBy,
                        "User who last modified the item"))
                .addPropertyList(property(MIME_TYPE, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, "MIME type when the item is a file"))
                .addPropertyList(property(SIZE, SourcePropertyType.Int64,
                        QUERYABLE_RETRIEVABLE, "Size in bytes; 0 for folders"))
                .addPropertyList(property(PARENT_ID, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.ItemParentId,
                        "Parent drive-item id"))
                .addPropertyList(property(DRIVE_ID, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, "Parent drive id"))
                .addPropertyList(property(FOLDER, SourcePropertyType.Boolean,
                        QUERYABLE_RETRIEVABLE, "True when the item is a folder"))
                .addPropertyList(property(ITEM_TYPE, SourcePropertyType.String,
                        QUERYABLE_RETRIEVABLE, SearchPropertyLabel.ItemType,
                        "file or folder"))
                .build();
    }

    private static SourcePropertyDefinition property(String name, SourcePropertyType type,
            int annotations, String description) {
        return SourcePropertyDefinition.newBuilder()
                .setName(name)
                .setType(type)
                .setDefaultSearchAnnotations(annotations)
                .setDescription(description)
                .build();
    }

    private static SourcePropertyDefinition property(String name, SourcePropertyType type,
            int annotations, SearchPropertyLabel label, String description) {
        return SourcePropertyDefinition.newBuilder()
                .setName(name)
                .setType(type)
                .setDefaultSearchAnnotations(annotations)
                .addDefaultSemanticLabels(label)
                .setDescription(description)
                .build();
    }
}
