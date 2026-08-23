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

    /** Display name of the drive item. */
    public static final String TITLE = "Title";
    /** File or folder name. */
    public static final String FILE_NAME = "FileName";
    /** File extension without the leading dot. */
    public static final String FILE_EXTENSION = "FileExtension";
    /** Browser URL of the item. */
    public static final String WEB_URL = "WebUrl";
    /** When the item was created. */
    public static final String CREATED = "Created";
    /** When the item was last modified. */
    public static final String LAST_MODIFIED = "LastModified";
    /** User who created the item. */
    public static final String CREATED_BY = "CreatedBy";
    /** User who last modified the item. */
    public static final String LAST_MODIFIED_BY = "LastModifiedBy";
    /** MIME type when the item is a file. */
    public static final String MIME_TYPE = "MimeType";
    /** Size in bytes; 0 for folders. */
    public static final String SIZE = "Size";
    /** Parent drive-item id. */
    public static final String PARENT_ID = "ParentId";
    /** Parent drive id. */
    public static final String DRIVE_ID = "DriveId";
    /** {@code true} when the item is a folder. */
    public static final String FOLDER = "Folder";
    /** {@code file} or {@code folder}. */
    public static final String ITEM_TYPE = "ItemType";

    private static final int SEARCHABLE_QUERYABLE_RETRIEVABLE =
            SearchAnnotations.IsSearchable_VALUE
                    | SearchAnnotations.IsQueryable_VALUE
                    | SearchAnnotations.IsRetrievable_VALUE;
    private static final int QUERYABLE_RETRIEVABLE =
            SearchAnnotations.IsQueryable_VALUE | SearchAnnotations.IsRetrievable_VALUE;

    private DataSourceSchemas() {
    }

    /**
     * The drive-item {@link DataSourceSchema} registered on the connection.
     *
     * @return the schema GCA stores
     */
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
