package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.Audio;
import ai.pipestream.microsoft.v1.BooleanColumn;
import ai.pipestream.microsoft.v1.CalculatedColumn;
import ai.pipestream.microsoft.v1.ChoiceColumn;
import ai.pipestream.microsoft.v1.ColumnDefinition;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.ColumnValidation;
import ai.pipestream.microsoft.v1.ContentApprovalStatusColumn;
import ai.pipestream.microsoft.v1.ContentTypeInfo;
import ai.pipestream.microsoft.v1.CurrencyColumn;
import ai.pipestream.microsoft.v1.DateTimeColumn;
import ai.pipestream.microsoft.v1.DefaultColumnValue;
import ai.pipestream.microsoft.v1.Deleted;
import ai.pipestream.microsoft.v1.DisplayNameLocalization;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemVersion;
import ai.pipestream.microsoft.v1.FileHashes;
import ai.pipestream.microsoft.v1.FileSystemInfo;
import ai.pipestream.microsoft.v1.GeoCoordinates;
import ai.pipestream.microsoft.v1.GeolocationColumn;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.HyperlinkOrPictureColumn;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.Image;
import ai.pipestream.microsoft.v1.ItemReference;
import ai.pipestream.microsoft.v1.ItemRetentionLabel;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.ListInfo;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.LookupColumn;
import ai.pipestream.microsoft.v1.Malware;
import ai.pipestream.microsoft.v1.NumberColumn;
import ai.pipestream.microsoft.v1.PackageFacet;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.PersonOrGroupColumn;
import ai.pipestream.microsoft.v1.Photo;
import ai.pipestream.microsoft.v1.PublicationFacet;
import ai.pipestream.microsoft.v1.Quota;
import ai.pipestream.microsoft.v1.RemoteItem;
import ai.pipestream.microsoft.v1.RetentionBehavior;
import ai.pipestream.microsoft.v1.RetentionLabelSettings;
import ai.pipestream.microsoft.v1.Shared;
import ai.pipestream.microsoft.v1.SharePointGroupIdentity;
import ai.pipestream.microsoft.v1.SharePointIdentity;
import ai.pipestream.microsoft.v1.SharePointIdentitySet;
import ai.pipestream.microsoft.v1.SharePointList;
import ai.pipestream.microsoft.v1.SharepointIds;
import ai.pipestream.microsoft.v1.SharingInvitation;
import ai.pipestream.microsoft.v1.SharingLink;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.microsoft.v1.SiteCollection;
import ai.pipestream.microsoft.v1.SpecialFolder;
import ai.pipestream.microsoft.v1.TermColumn;
import ai.pipestream.microsoft.v1.TextColumn;
import ai.pipestream.microsoft.v1.ThumbnailColumn;
import ai.pipestream.microsoft.v1.Video;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.Timestamp;

import java.time.Instant;
import java.time.OffsetDateTime;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * Jackson {@link JsonNode} to domain proto translation for the Microsoft Graph
 * v1.0 SharePoint / OneDrive resources the contract models: site, drive,
 * driveItem (every facet), list, listItem, columnDefinition, permission,
 * driveItemVersion, itemRetentionLabel and user. Every modelled property is
 * mapped; a property Graph omitted stays unset (facets keep Graph's
 * presence-means-true semantics through message presence). Unknown values
 * never throw. The deprecated permission {@code grantedTo} /
 * {@code grantedToIdentities} are folded into the V2 fields when Graph sent
 * no V2 value.
 */
public final class MicrosoftMapper {

    /** Graph reports "never expires" as DateTime.MinValue; it is dropped. */
    private static final long DATE_TIME_MIN_EPOCH_SECONDS = -62135596800L;

    /** Creates a mapper. */
    public MicrosoftMapper() {
    }

    // ------------------------------------------------------------------ users

    /**
     * Maps a Graph user JSON object to the domain proto.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public GraphUser toUser(JsonNode node) {
        GraphUser.Builder b = GraphUser.newBuilder()
                .setId(text(node, "id"))
                .setDisplayName(text(node, "displayName"))
                .setUserPrincipalName(text(node, "userPrincipalName"))
                .setMail(text(node, "mail"))
                .setGivenName(text(node, "givenName"))
                .setSurname(text(node, "surname"))
                .setJobTitle(text(node, "jobTitle"))
                .setDepartment(text(node, "department"))
                .setOfficeLocation(text(node, "officeLocation"))
                .setUserType(text(node, "userType"))
                .setPreferredLanguage(text(node, "preferredLanguage"));
        if (node.path("accountEnabled").isBoolean()) {
            b.setAccountEnabled(node.path("accountEnabled").booleanValue());
        }
        return b.build();
    }

    // ------------------------------------------------------------------ sites

    /**
     * Maps a Graph site JSON object to the domain proto. A {@code deleted} facet
     * or a {@code @removed} annotation (the generic delta-query delete marker)
     * becomes {@link Site#getDeleted()}.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public Site toSite(JsonNode node) {
        Site.Builder b = Site.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDisplayName(text(node, "displayName"))
                .setWebUrl(text(node, "webUrl"))
                .setDescription(text(node, "description"))
                .setEtag(firstText(node, "eTag", "etag"))
                .setIsPersonalSite(node.path("isPersonalSite").asBoolean(false))
                .setIsRoot(node.path("root").isObject());
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        JsonNode collection = node.path("siteCollection");
        if (collection.isObject()) {
            b.setSiteCollection(SiteCollection.newBuilder()
                    .setHostname(firstText(collection, "hostname", "hostName"))
                    .setDataLocationCode(text(collection, "dataLocationCode"))
                    .setIsRoot(collection.path("root").isObject())
                    .setArchivalState(text(collection.path("archivalDetails"), "archiveStatus"))
                    .build());
        }
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        itemReference(node.path("parentReference"), b::setParentReference);
        deleted(node, b::setDeleted);
        return b.build();
    }

    // ------------------------------------------------------------------ drives

    /**
     * Maps a Graph drive JSON object to the domain proto.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public Drive toDrive(JsonNode node) {
        Drive.Builder b = Drive.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDriveType(text(node, "driveType"))
                .setWebUrl(text(node, "webUrl"))
                .setDescription(text(node, "description"))
                .setEtag(firstText(node, "eTag", "etag"))
                .setIsSystem(node.path("system").isObject());
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        identitySet(node, "owner", b::setOwner);
        JsonNode quota = node.path("quota");
        if (quota.isObject()) {
            b.setQuota(Quota.newBuilder()
                    .setTotal(longValue(quota, "total"))
                    .setUsed(longValue(quota, "used"))
                    .setRemaining(longValue(quota, "remaining"))
                    .setDeleted(longValue(quota, "deleted"))
                    .setState(text(quota, "state"))
                    .build());
        }
        JsonNode ids = node.has("sharePointIds") ? node.path("sharePointIds")
                : node.path("sharepointIds");
        sharepointIds(ids, b::setSharepointIds);
        b.setListId(text(ids, "listId"));
        return b.build();
    }

    /**
     * Maps a Graph drive JSON object and sets {@code siteId} when non-blank.
     *
     * @param node the API object
     * @param siteId parent SharePoint site id; blank leaves the field unset
     * @return the proto
     */
    public Drive toDrive(JsonNode node, String siteId) {
        Drive.Builder b = toDrive(node).toBuilder();
        if (siteId != null && !siteId.isBlank()) {
            b.setSiteId(siteId);
        }
        return b.build();
    }

    // -------------------------------------------------------------- drive items

    /**
     * Maps a Graph driveItem JSON object to the domain proto, every facet included.
     *
     * @param node the API object
     * @param driveId parent drive id when the JSON omits {@code parentReference}
     * @return the proto; missing fields stay unset
     */
    public DriveItem toDriveItem(JsonNode node, String driveId) {
        DriveItem.Builder b = DriveItem.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDriveId(driveId == null ? "" : driveId)
                .setWebUrl(text(node, "webUrl"))
                .setDownloadUrl(firstText(node, "@microsoft.graph.downloadUrl", "@content.downloadUrl"))
                .setSize(longValue(node, "size"))
                .setFolder(node.path("folder").isObject())
                .setDescription(text(node, "description"))
                .setEtag(firstText(node, "eTag", "etag"))
                .setCtag(firstText(node, "cTag", "ctag"))
                .setWebDavUrl(text(node, "webDavUrl"))
                .setIsRoot(node.path("root").isObject())
                .setPendingContentUpdate(node.path("pendingOperations")
                        .path("pendingContentUpdate").isObject())
                .setSharingChanged(truthy(node.path("@microsoft.graph.sharedChanged")));
        JsonNode folder = node.path("folder");
        if (folder.isObject() && folder.path("childCount").isIntegralNumber()) {
            b.setChildCount(folder.path("childCount").intValue());
        }
        JsonNode file = node.path("file");
        if (file.isObject()) {
            b.setMimeType(text(file, "mimeType"));
            hashes(file.path("hashes"), b::setHashes);
        }
        JsonNode parent = node.path("parentReference");
        if (parent.isObject()) {
            b.setParentId(text(parent, "id"));
            if (b.getDriveId().isEmpty()) {
                b.setDriveId(text(parent, "driveId"));
            }
            b.setParentReference(toItemReference(parent));
        }
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        fileSystemInfo(node.path("fileSystemInfo"), b::setFileSystemInfo);
        deleted(node, b::setDeleted);
        if (node.path("package").isObject()) {
            b.setPackage(PackageFacet.newBuilder().setType(text(node.path("package"), "type")));
        }
        if (node.path("remoteItem").isObject()) {
            b.setRemoteItem(toRemoteItem(node.path("remoteItem")));
        }
        shared(node.path("shared"), b::setShared);
        specialFolder(node.path("specialFolder"), b::setSpecialFolder);
        image(node.path("image"), b::setImage);
        if (node.path("photo").isObject()) {
            b.setPhoto(toPhoto(node.path("photo")));
        }
        video(node.path("video"), b::setVideo);
        if (node.path("audio").isObject()) {
            b.setAudio(toAudio(node.path("audio")));
        }
        if (node.path("location").isObject()) {
            JsonNode loc = node.path("location");
            GeoCoordinates.Builder geo = GeoCoordinates.newBuilder();
            if (loc.path("altitude").isNumber()) {
                geo.setAltitude(loc.path("altitude").doubleValue());
            }
            if (loc.path("latitude").isNumber()) {
                geo.setLatitude(loc.path("latitude").doubleValue());
            }
            if (loc.path("longitude").isNumber()) {
                geo.setLongitude(loc.path("longitude").doubleValue());
            }
            b.setLocation(geo);
        }
        if (node.path("malware").isObject()) {
            b.setMalware(Malware.newBuilder().setDescription(text(node.path("malware"), "description")));
        }
        publication(node.path("publication"), b::setPublication);
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        JsonNode listItem = node.path("listItem");
        if (listItem.isObject()) {
            b.setListItemId(text(listItem, "id"));
            contentType(listItem.path("contentType"), b::setContentType);
            if (listItem.path("fields").isObject()) {
                b.addAllListColumns(toListColumns(listItem.path("fields")));
            }
        } else if (b.hasSharepointIds()) {
            b.setListItemId(b.getSharepointIds().getListItemId());
        }
        return b.build();
    }

    /**
     * Copies {@code item} and replaces {@code list_columns} with the flattened
     * SharePoint fields object. {@code @odata.*} keys are skipped; nested
     * objects flatten as {@code Parent.Child}.
     *
     * @param item the drive item
     * @param fields Graph {@code listItem.fields} object; {@code null} or
     *        non-object leaves columns empty
     * @return a copy with typed columns
     */
    public DriveItem withListColumns(DriveItem item, JsonNode fields) {
        return item.toBuilder().clearListColumns().addAllListColumns(toListColumns(fields)).build();
    }

    /**
     * Copies {@code item} with the backing list item applied: its id, content
     * type and typed {@code fields}.
     *
     * @param item the drive item
     * @param listItem Graph {@code listItem} JSON (with {@code fields} expanded);
     *        {@code null} or non-object leaves the item unchanged except for
     *        clearing columns
     * @return a copy carrying the list-item view
     */
    public DriveItem withListItem(DriveItem item, JsonNode listItem) {
        DriveItem.Builder b = item.toBuilder().clearListColumns();
        if (listItem == null || !listItem.isObject()) {
            return b.build();
        }
        if (!text(listItem, "id").isEmpty()) {
            b.setListItemId(text(listItem, "id"));
        }
        contentType(listItem.path("contentType"), b::setContentType);
        b.addAllListColumns(toListColumns(listItem.path("fields")));
        return b.build();
    }

    /**
     * Maps the {@code remoteItem} facet.
     *
     * @param node the facet object
     * @return the proto
     */
    public RemoteItem toRemoteItem(JsonNode node) {
        RemoteItem.Builder b = RemoteItem.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setWebUrl(text(node, "webUrl"))
                .setWebDavUrl(text(node, "webDavUrl"))
                .setSize(longValue(node, "size"))
                .setFolder(node.path("folder").isObject());
        itemReference(node.path("parentReference"), b::setParentReference);
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        if (node.path("folder").path("childCount").isIntegralNumber()) {
            b.setChildCount(node.path("folder").path("childCount").intValue());
        }
        JsonNode file = node.path("file");
        if (file.isObject()) {
            b.setMimeType(text(file, "mimeType"));
            hashes(file.path("hashes"), b::setHashes);
        }
        fileSystemInfo(node.path("fileSystemInfo"), b::setFileSystemInfo);
        if (node.path("package").isObject()) {
            b.setPackage(PackageFacet.newBuilder().setType(text(node.path("package"), "type")));
        }
        shared(node.path("shared"), b::setShared);
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        specialFolder(node.path("specialFolder"), b::setSpecialFolder);
        image(node.path("image"), b::setImage);
        video(node.path("video"), b::setVideo);
        return b.build();
    }

    /**
     * Maps a {@code driveItemVersion}.
     *
     * @param node the API object
     * @param driveId owning drive, set on the result
     * @param itemId owning item, set on the result
     * @return the proto
     */
    public DriveItemVersion toVersion(JsonNode node, String driveId, String itemId) {
        DriveItemVersion.Builder b = DriveItemVersion.newBuilder()
                .setId(text(node, "id"))
                .setSize(longValue(node, "size"))
                .setDriveId(driveId)
                .setItemId(itemId);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        publication(node.path("publication"), b::setPublication);
        return b.build();
    }

    /**
     * Maps an {@code itemRetentionLabel}.
     *
     * @param node the API object
     * @return the proto
     */
    public ItemRetentionLabel toRetentionLabel(JsonNode node) {
        ItemRetentionLabel.Builder b = ItemRetentionLabel.newBuilder()
                .setName(text(node, "name"))
                .setIsLabelAppliedExplicitly(node.path("isLabelAppliedExplicitly").asBoolean(false));
        identitySet(node, "labelAppliedBy", b::setLabelAppliedBy);
        ts(node, "labelAppliedDateTime", b::setLabelAppliedAt);
        JsonNode settings = node.path("retentionSettings");
        if (settings.isObject()) {
            b.setRetentionSettings(RetentionLabelSettings.newBuilder()
                    .setBehaviorDuringRetentionPeriod(retentionBehavior(
                            text(settings, "behaviorDuringRetentionPeriod")))
                    .setIsContentUpdateAllowed(settings.path("isContentUpdateAllowed").asBoolean(false))
                    .setIsDeleteAllowed(settings.path("isDeleteAllowed").asBoolean(false))
                    .setIsLabelUpdateAllowed(settings.path("isLabelUpdateAllowed").asBoolean(false))
                    .setIsMetadataUpdateAllowed(settings.path("isMetadataUpdateAllowed").asBoolean(false))
                    .setIsRecordLocked(settings.path("isRecordLocked").asBoolean(false))
                    .build());
        }
        return b.build();
    }

    // ------------------------------------------------------------------ lists

    /**
     * Maps a Graph {@code list}.
     *
     * @param node the API object
     * @param siteId owning site; blank falls back to {@code parentReference.siteId}
     * @return the proto
     */
    public SharePointList toList(JsonNode node, String siteId) {
        SharePointList.Builder b = SharePointList.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDisplayName(text(node, "displayName"))
                .setDescription(text(node, "description"))
                .setWebUrl(text(node, "webUrl"))
                .setEtag(firstText(node, "eTag", "etag"))
                .setIsSystem(node.path("system").isObject())
                .setDriveId(text(node.path("drive"), "id"));
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        JsonNode info = node.path("list");
        if (info.isObject()) {
            b.setList(ListInfo.newBuilder()
                    .setContentTypesEnabled(info.path("contentTypesEnabled").asBoolean(false))
                    .setHidden(info.path("hidden").asBoolean(false))
                    .setTemplate(text(info, "template")));
        }
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        itemReference(node.path("parentReference"), b::setParentReference);
        String site = siteId == null || siteId.isBlank()
                ? text(node.path("parentReference"), "siteId") : siteId;
        b.setSiteId(site);
        return b.build();
    }

    /**
     * Maps a Graph {@code listItem} with its typed {@code fields}.
     *
     * @param node the API object
     * @param siteId owning site, set on the result
     * @param listId owning list, set on the result
     * @return the proto
     */
    public ListItem toListItem(JsonNode node, String siteId, String listId) {
        ListItem.Builder b = ListItem.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setWebUrl(text(node, "webUrl"))
                .setDescription(text(node, "description"))
                .setEtag(firstText(node, "eTag", "etag"))
                .setSiteId(siteId == null ? "" : siteId)
                .setListId(listId == null ? "" : listId);
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        identitySet(node, "createdBy", b::setCreatedBy);
        identitySet(node, "lastModifiedBy", b::setLastModifiedBy);
        contentType(node.path("contentType"), b::setContentType);
        b.addAllFields(toListColumns(node.path("fields")));
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        deleted(node, b::setDeleted);
        itemReference(node.path("parentReference"), b::setParentReference);
        JsonNode driveItem = node.path("driveItem");
        if (driveItem.isObject()) {
            b.setDriveItem(toDriveItem(driveItem, text(driveItem.path("parentReference"), "driveId")));
        }
        return b.build();
    }

    /**
     * Maps a Graph {@code columnDefinition}.
     *
     * @param node the API object
     * @param siteId site the column was read from
     * @param listId list the column was read from; empty for site / content type
     * @param contentTypeId content type the column was read from; empty otherwise
     * @return the proto
     */
    public ColumnDefinition toColumn(JsonNode node, String siteId, String listId,
            String contentTypeId) {
        ColumnDefinition.Builder b = ColumnDefinition.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDisplayName(text(node, "displayName"))
                .setDescription(text(node, "description"))
                .setColumnGroup(text(node, "columnGroup"))
                .setType(columnType(text(node, "type")))
                .setHidden(node.path("hidden").asBoolean(false))
                .setIndexed(node.path("indexed").asBoolean(false))
                .setReadOnly(node.path("readOnly").asBoolean(false))
                .setRequired(node.path("required").asBoolean(false))
                .setEnforceUniqueValues(node.path("enforceUniqueValues").asBoolean(false))
                .setIsDeletable(node.path("isDeletable").asBoolean(false))
                .setIsReorderable(node.path("isReorderable").asBoolean(false))
                .setIsSealed(node.path("isSealed").asBoolean(false))
                .setPropagateChanges(node.path("propagateChanges").asBoolean(false))
                .setSiteId(siteId == null ? "" : siteId)
                .setListId(listId == null ? "" : listId)
                .setContentTypeId(contentTypeId == null ? "" : contentTypeId);
        JsonNode defaultValue = node.path("defaultValue");
        if (defaultValue.isObject()) {
            b.setDefaultValue(DefaultColumnValue.newBuilder()
                    .setFormula(text(defaultValue, "formula"))
                    .setValue(text(defaultValue, "value")));
        }
        JsonNode validation = node.path("validation");
        if (validation.isObject()) {
            ColumnValidation.Builder v = ColumnValidation.newBuilder()
                    .setDefaultLanguage(text(validation, "defaultLanguage"))
                    .setFormula(text(validation, "formula"));
            for (JsonNode d : validation.path("descriptions")) {
                v.addDescriptions(DisplayNameLocalization.newBuilder()
                        .setDisplayName(text(d, "displayName"))
                        .setLanguageTag(text(d, "languageTag")));
            }
            b.setValidation(v);
        }
        contentType(node.path("sourceContentType"), b::setSourceContentType);
        columnKind(node, b);
        return b.build();
    }

    private static void columnKind(JsonNode node, ColumnDefinition.Builder b) {
        JsonNode f;
        if ((f = node.path("text")).isObject()) {
            b.setText(TextColumn.newBuilder()
                    .setAllowMultipleLines(f.path("allowMultipleLines").asBoolean(false))
                    .setAppendChangesToExistingText(f.path("appendChangesToExistingText").asBoolean(false))
                    .setLinesForEditing(f.path("linesForEditing").asInt(0))
                    .setMaxLength(f.path("maxLength").asInt(0))
                    .setTextType(text(f, "textType")));
        } else if ((f = node.path("choice")).isObject()) {
            ChoiceColumn.Builder c = ChoiceColumn.newBuilder()
                    .setAllowTextEntry(f.path("allowTextEntry").asBoolean(false))
                    .setDisplayAs(text(f, "displayAs"));
            for (JsonNode choice : f.path("choices")) {
                c.addChoices(choice.asText());
            }
            b.setChoice(c);
        } else if ((f = node.path("lookup")).isObject()) {
            b.setLookup(LookupColumn.newBuilder()
                    .setAllowMultipleValues(f.path("allowMultipleValues").asBoolean(false))
                    .setAllowUnlimitedLength(f.path("allowUnlimitedLength").asBoolean(false))
                    .setColumnName(text(f, "columnName"))
                    .setListId(text(f, "listId"))
                    .setPrimaryLookupColumnId(text(f, "primaryLookupColumnId")));
        } else if ((f = node.path("personOrGroup")).isObject()) {
            b.setPersonOrGroup(PersonOrGroupColumn.newBuilder()
                    .setAllowMultipleSelection(f.path("allowMultipleSelection").asBoolean(false))
                    .setChooseFromType(text(f, "chooseFromType"))
                    .setDisplayAs(text(f, "displayAs")));
        } else if ((f = node.path("dateTime")).isObject()) {
            b.setDateTime(DateTimeColumn.newBuilder()
                    .setDisplayAs(text(f, "displayAs"))
                    .setFormat(text(f, "format")));
        } else if ((f = node.path("number")).isObject()) {
            NumberColumn.Builder n = NumberColumn.newBuilder()
                    .setDecimalPlaces(text(f, "decimalPlaces"))
                    .setDisplayAs(text(f, "displayAs"));
            if (f.path("maximum").isNumber()) {
                n.setMaximum(f.path("maximum").doubleValue());
            }
            if (f.path("minimum").isNumber()) {
                n.setMinimum(f.path("minimum").doubleValue());
            }
            b.setNumber(n);
        } else if ((f = node.path("currency")).isObject()) {
            b.setCurrency(CurrencyColumn.newBuilder().setLocale(text(f, "locale")));
        } else if ((f = node.path("calculated")).isObject()) {
            b.setCalculated(CalculatedColumn.newBuilder()
                    .setFormat(text(f, "format"))
                    .setFormula(text(f, "formula"))
                    .setOutputType(text(f, "outputType")));
        } else if ((f = node.path("hyperlinkOrPicture")).isObject()) {
            b.setHyperlinkOrPicture(HyperlinkOrPictureColumn.newBuilder()
                    .setIsPicture(f.path("isPicture").asBoolean(false)));
        } else if ((f = node.path("term")).isObject()) {
            b.setTerm(TermColumn.newBuilder()
                    .setAllowMultipleValues(f.path("allowMultipleValues").asBoolean(false))
                    .setShowFullyQualifiedName(f.path("showFullyQualifiedName").asBoolean(false))
                    .setTermSetId(text(f.path("termSet"), "id"))
                    .setParentTermId(text(f.path("parentTerm"), "id")));
        } else if (node.path("boolean").isObject()) {
            b.setBoolean(BooleanColumn.getDefaultInstance());
        } else if (node.path("geolocation").isObject()) {
            b.setGeolocation(GeolocationColumn.getDefaultInstance());
        } else if (node.path("thumbnail").isObject()) {
            b.setThumbnail(ThumbnailColumn.getDefaultInstance());
        } else if (node.path("contentApprovalStatus").isObject()) {
            b.setContentApprovalStatus(ContentApprovalStatusColumn.getDefaultInstance());
        }
    }

    // ------------------------------------------------------------- permissions

    /**
     * Maps a Graph {@code permission}. Deprecated {@code grantedTo} /
     * {@code grantedToIdentities} fill the V2 fields when Graph sent no V2 value.
     *
     * @param node the API object
     * @param target the resource the permission is on; {@code null} leaves it unset
     * @return the proto
     */
    public Permission toPermission(JsonNode node, PermissionTarget target) {
        Permission.Builder b = Permission.newBuilder()
                .setId(text(node, "id"))
                .setShareId(text(node, "shareId"))
                .setHasPassword(node.path("hasPassword").asBoolean(false));
        for (JsonNode role : node.path("roles")) {
            b.addRoles(role.asText());
        }
        JsonNode grantedTo = node.path("grantedToV2").isObject() ? node.path("grantedToV2")
                : node.path("grantedTo");
        if (grantedTo.isObject()) {
            b.setGrantedTo(toSharePointIdentitySet(grantedTo));
        }
        JsonNode identities = node.path("grantedToIdentitiesV2").isArray()
                && !node.path("grantedToIdentitiesV2").isEmpty()
                ? node.path("grantedToIdentitiesV2") : node.path("grantedToIdentities");
        for (JsonNode identity : identities) {
            if (identity.isObject()) {
                b.addGrantedToIdentities(toSharePointIdentitySet(identity));
            }
        }
        JsonNode link = node.path("link");
        if (link.isObject()) {
            SharingLink.Builder l = SharingLink.newBuilder()
                    .setType(text(link, "type"))
                    .setScope(text(link, "scope"))
                    .setWebUrl(text(link, "webUrl"))
                    .setWebHtml(text(link, "webHtml"))
                    .setPreventsDownload(link.path("preventsDownload").asBoolean(false));
            if (link.path("application").isObject()) {
                l.setApplication(toIdentity(link.path("application")));
            }
            b.setLink(l);
        }
        JsonNode invitation = node.path("invitation");
        if (invitation.isObject()) {
            SharingInvitation.Builder i = SharingInvitation.newBuilder()
                    .setEmail(text(invitation, "email"))
                    .setRedeemedBy(text(invitation, "redeemedBy"))
                    .setSignInRequired(invitation.path("signInRequired").asBoolean(false));
            identitySet(invitation, "invitedBy", i::setInvitedBy);
            b.setInvitation(i);
        }
        itemReference(node.path("inheritedFrom"), b::setInheritedFrom);
        Timestamp expires = timestamp(text(node, "expirationDateTime"));
        if (expires != null && expires.getSeconds() > DATE_TIME_MIN_EPOCH_SECONDS) {
            b.setExpiresAt(expires);
        }
        if (target != null) {
            b.setTarget(target);
        }
        return b.build();
    }

    private static SharePointIdentitySet toSharePointIdentitySet(JsonNode node) {
        SharePointIdentitySet.Builder b = SharePointIdentitySet.newBuilder();
        if (node.path("user").isObject()) {
            b.setUser(toIdentity(node.path("user")));
        }
        if (node.path("application").isObject()) {
            b.setApplication(toIdentity(node.path("application")));
        }
        if (node.path("device").isObject()) {
            b.setDevice(toIdentity(node.path("device")));
        }
        if (node.path("group").isObject()) {
            b.setGroup(toIdentity(node.path("group")));
        }
        if (node.path("siteUser").isObject()) {
            b.setSiteUser(toSharePointIdentity(node.path("siteUser")));
        }
        if (node.path("siteGroup").isObject()) {
            b.setSiteGroup(toSharePointIdentity(node.path("siteGroup")));
        }
        JsonNode group = node.path("sharePointGroup");
        if (group.isObject()) {
            b.setSharePointGroup(SharePointGroupIdentity.newBuilder()
                    .setId(text(group, "id"))
                    .setDisplayName(text(group, "displayName"))
                    .setPrincipalId(text(group, "principalId"))
                    .setTitle(text(group, "title")));
        }
        return b.build();
    }

    private static SharePointIdentity toSharePointIdentity(JsonNode node) {
        return SharePointIdentity.newBuilder()
                .setId(text(node, "id"))
                .setDisplayName(text(node, "displayName"))
                .setLoginName(text(node, "loginName"))
                .build();
    }

    // ---------------------------------------------------------------- columns

    /**
     * Flattens a SharePoint {@code fields} object into typed {@link ListColumn}s.
     * Objects flatten into dotted names at any depth; an array of scalars is one
     * column with {@code string_values} (and the joined {@code string_value});
     * an array of objects becomes one column per property path whose
     * {@code string_values} hold that property of each element.
     *
     * @param fields Graph fields object; {@code null} or non-object yields empty
     * @return columns in field-name order
     */
    public List<ListColumn> toListColumns(JsonNode fields) {
        List<ListColumn> columns = new ArrayList<>();
        if (fields == null || !fields.isObject()) {
            return columns;
        }
        Iterator<Map.Entry<String, JsonNode>> it = fields.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> entry = it.next();
            addColumns(columns, entry.getKey(), entry.getValue());
        }
        return columns;
    }

    private static void addColumns(List<ListColumn> out, String name, JsonNode value) {
        if (name == null || name.startsWith("@")) {
            return;
        }
        if (value == null || value.isNull() || value.isMissingNode()) {
            return;
        }
        if (value.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> nested = value.fields();
            while (nested.hasNext()) {
                Map.Entry<String, JsonNode> entry = nested.next();
                if (!entry.getKey().startsWith("@")) {
                    addColumns(out, name + "." + entry.getKey(), entry.getValue());
                }
            }
            return;
        }
        if (value.isArray()) {
            addArrayColumns(out, name, value);
            return;
        }
        out.add(scalarColumn(name, value));
    }

    private static void addArrayColumns(List<ListColumn> out, String name, JsonNode array) {
        List<String> scalars = new ArrayList<>();
        Map<String, List<String>> objectPaths = new LinkedHashMap<>();
        for (JsonNode element : array) {
            if (element.isNull()) {
                continue;
            }
            if (element.isValueNode()) {
                scalars.add(element.asText());
            } else if (element.isObject()) {
                Map<String, String> leaves = new LinkedHashMap<>();
                leaves(element, "", leaves);
                leaves.forEach((path, text) ->
                        objectPaths.computeIfAbsent(path, k -> new ArrayList<>()).add(text));
            } else if (element.isArray()) {
                for (JsonNode inner : element) {
                    if (inner.isValueNode() && !inner.isNull()) {
                        scalars.add(inner.asText());
                    }
                }
            }
        }
        if (!scalars.isEmpty()) {
            out.add(ListColumn.newBuilder().setName(name)
                    .setStringValue(String.join(", ", scalars))
                    .addAllStringValues(scalars).build());
        }
        objectPaths.forEach((path, values) -> out.add(ListColumn.newBuilder()
                .setName(name + "." + path)
                .setStringValue(String.join(", ", values))
                .addAllStringValues(values).build()));
    }

    private static void leaves(JsonNode node, String prefix, Map<String, String> out) {
        Iterator<Map.Entry<String, JsonNode>> it = node.fields();
        while (it.hasNext()) {
            Map.Entry<String, JsonNode> entry = it.next();
            if (entry.getKey().startsWith("@")) {
                continue;
            }
            String path = prefix.isEmpty() ? entry.getKey() : prefix + "." + entry.getKey();
            JsonNode value = entry.getValue();
            if (value.isObject()) {
                leaves(value, path, out);
            } else if (value.isArray()) {
                List<String> parts = new ArrayList<>();
                value.forEach(v -> {
                    if (v.isValueNode() && !v.isNull()) {
                        parts.add(v.asText());
                    }
                });
                if (!parts.isEmpty()) {
                    out.put(path, String.join(", ", parts));
                }
            } else if (!value.isNull()) {
                out.put(path, value.asText());
            }
        }
    }

    private static ListColumn scalarColumn(String name, JsonNode value) {
        ListColumn.Builder b = ListColumn.newBuilder().setName(name);
        if (value.isBoolean()) {
            return b.setBoolValue(value.booleanValue()).build();
        }
        if (value.isIntegralNumber() && value.canConvertToLong()) {
            return b.setIntValue(value.longValue()).build();
        }
        if (value.isNumber()) {
            return b.setDoubleValue(value.doubleValue()).build();
        }
        String text = value.asText();
        if (value.isTextual() && looksLikeTimestamp(text)) {
            Timestamp ts = timestamp(text);
            if (ts != null) {
                return b.setTimestampValue(ts).build();
            }
        }
        return b.setStringValue(text).build();
    }

    private static boolean looksLikeTimestamp(String text) {
        return text.length() > 10 && text.charAt(4) == '-' && text.charAt(10) == 'T';
    }

    // ---------------------------------------------------------------- facets

    /**
     * Maps an {@code itemReference}.
     *
     * @param node the API object
     * @return the proto
     */
    public ItemReference toItemReference(JsonNode node) {
        ItemReference.Builder b = ItemReference.newBuilder()
                .setDriveId(text(node, "driveId"))
                .setDriveType(text(node, "driveType"))
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setPath(text(node, "path"))
                .setShareId(text(node, "shareId"))
                .setSiteId(text(node, "siteId"));
        sharepointIds(node.path("sharepointIds"), b::setSharepointIds);
        return b.build();
    }

    private void itemReference(JsonNode node, Consumer<ItemReference> setter) {
        if (node.isObject()) {
            setter.accept(toItemReference(node));
        }
    }

    private static void sharepointIds(JsonNode node, Consumer<SharepointIds> setter) {
        if (!node.isObject()) {
            return;
        }
        setter.accept(SharepointIds.newBuilder()
                .setListId(text(node, "listId"))
                .setListItemId(text(node, "listItemId"))
                .setListItemUniqueId(text(node, "listItemUniqueId"))
                .setSiteId(text(node, "siteId"))
                .setSiteUrl(text(node, "siteUrl"))
                .setTenantId(text(node, "tenantId"))
                .setWebId(text(node, "webId"))
                .build());
    }

    private static void hashes(JsonNode node, Consumer<FileHashes> setter) {
        if (!node.isObject()) {
            return;
        }
        setter.accept(FileHashes.newBuilder()
                .setSha1(firstText(node, "sha1Hash", "sha1"))
                .setSha256(firstText(node, "sha256Hash", "sha256"))
                .setQuickXor(firstText(node, "quickXorHash", "quickXor"))
                .setCrc32(firstText(node, "crc32Hash", "crc32"))
                .build());
    }

    private static void fileSystemInfo(JsonNode node, Consumer<FileSystemInfo> setter) {
        if (!node.isObject()) {
            return;
        }
        FileSystemInfo.Builder b = FileSystemInfo.newBuilder();
        ts(node, "createdDateTime", b::setCreatedAt);
        ts(node, "lastModifiedDateTime", b::setLastModifiedAt);
        ts(node, "lastAccessedDateTime", b::setLastAccessedAt);
        setter.accept(b.build());
    }

    private static void deleted(JsonNode node, Consumer<Deleted> setter) {
        JsonNode deleted = node.path("deleted");
        if (deleted.isObject()) {
            setter.accept(Deleted.newBuilder().setState(text(deleted, "state")).build());
            return;
        }
        JsonNode removed = node.path("@removed");
        if (removed.isObject()) {
            setter.accept(Deleted.newBuilder().setState(text(removed, "reason")).build());
        }
    }

    private static void shared(JsonNode node, Consumer<Shared> setter) {
        if (!node.isObject()) {
            return;
        }
        Shared.Builder b = Shared.newBuilder().setScope(text(node, "scope"));
        identitySet(node, "owner", b::setOwner);
        identitySet(node, "sharedBy", b::setSharedBy);
        ts(node, "sharedDateTime", b::setSharedAt);
        setter.accept(b.build());
    }

    private static void specialFolder(JsonNode node, Consumer<SpecialFolder> setter) {
        if (node.isObject()) {
            setter.accept(SpecialFolder.newBuilder().setName(text(node, "name")).build());
        }
    }

    private static void image(JsonNode node, Consumer<Image> setter) {
        if (node.isObject()) {
            setter.accept(Image.newBuilder()
                    .setWidth(node.path("width").asInt(0))
                    .setHeight(node.path("height").asInt(0))
                    .build());
        }
    }

    private static void video(JsonNode node, Consumer<Video> setter) {
        if (!node.isObject()) {
            return;
        }
        Video.Builder b = Video.newBuilder()
                .setAudioBitsPerSample(node.path("audioBitsPerSample").asInt(0))
                .setAudioChannels(node.path("audioChannels").asInt(0))
                .setAudioFormat(text(node, "audioFormat"))
                .setAudioSamplesPerSecond(node.path("audioSamplesPerSecond").asInt(0))
                .setBitrate(node.path("bitrate").asInt(0))
                .setDurationMs(node.path("duration").asLong(0))
                .setFourCc(text(node, "fourCC"))
                .setHeight(node.path("height").asInt(0))
                .setWidth(node.path("width").asInt(0));
        if (node.path("frameRate").isNumber()) {
            b.setFrameRate(node.path("frameRate").doubleValue());
        }
        setter.accept(b.build());
    }

    private static Photo toPhoto(JsonNode node) {
        Photo.Builder b = Photo.newBuilder()
                .setCameraMake(text(node, "cameraMake"))
                .setCameraModel(text(node, "cameraModel"));
        if (node.path("exposureNumerator").isNumber()) {
            b.setExposureNumerator(node.path("exposureNumerator").doubleValue());
        }
        if (node.path("exposureDenominator").isNumber()) {
            b.setExposureDenominator(node.path("exposureDenominator").doubleValue());
        }
        if (node.path("fNumber").isNumber()) {
            b.setFNumber(node.path("fNumber").doubleValue());
        }
        if (node.path("focalLength").isNumber()) {
            b.setFocalLength(node.path("focalLength").doubleValue());
        }
        if (node.path("iso").isNumber()) {
            b.setIso(node.path("iso").intValue());
        }
        if (node.path("orientation").isNumber()) {
            b.setOrientation(node.path("orientation").intValue());
        }
        ts(node, "takenDateTime", b::setTakenAt);
        return b.build();
    }

    private static Audio toAudio(JsonNode node) {
        return Audio.newBuilder()
                .setAlbum(text(node, "album"))
                .setAlbumArtist(text(node, "albumArtist"))
                .setArtist(text(node, "artist"))
                .setBitrate(node.path("bitrate").asLong(0))
                .setComposers(text(node, "composers"))
                .setCopyright(text(node, "copyright"))
                .setDisc(node.path("disc").asInt(0))
                .setDiscCount(node.path("discCount").asInt(0))
                .setDurationMs(node.path("duration").asLong(0))
                .setGenre(text(node, "genre"))
                .setHasDrm(node.path("hasDrm").asBoolean(false))
                .setIsVariableBitrate(node.path("isVariableBitrate").asBoolean(false))
                .setTitle(text(node, "title"))
                .setTrack(node.path("track").asInt(0))
                .setTrackCount(node.path("trackCount").asInt(0))
                .setYear(node.path("year").asInt(0))
                .build();
    }

    private static void publication(JsonNode node, Consumer<PublicationFacet> setter) {
        if (!node.isObject()) {
            return;
        }
        PublicationFacet.Builder b = PublicationFacet.newBuilder()
                .setLevel(text(node, "level"))
                .setVersionId(text(node, "versionId"));
        identitySet(node, "checkedOutBy", b::setCheckedOutBy);
        setter.accept(b.build());
    }

    private static void contentType(JsonNode node, Consumer<ContentTypeInfo> setter) {
        if (node.isObject()) {
            setter.accept(ContentTypeInfo.newBuilder()
                    .setId(text(node, "id"))
                    .setName(text(node, "name"))
                    .build());
        }
    }

    private static void identitySet(JsonNode parent, String field, Consumer<IdentitySet> setter) {
        JsonNode node = parent.path(field);
        if (!node.isObject()) {
            return;
        }
        IdentitySet.Builder b = IdentitySet.newBuilder();
        if (node.path("user").isObject()) {
            b.setUser(toIdentity(node.path("user")));
        }
        if (node.path("application").isObject()) {
            b.setApplication(toIdentity(node.path("application")));
        }
        if (node.path("device").isObject()) {
            b.setDevice(toIdentity(node.path("device")));
        }
        if (node.path("group").isObject()) {
            b.setGroup(toIdentity(node.path("group")));
        }
        setter.accept(b.build());
    }

    private static Identity toIdentity(JsonNode node) {
        return Identity.newBuilder()
                .setId(text(node, "id"))
                .setDisplayName(text(node, "displayName"))
                .setEmail(text(node, "email"))
                .build();
    }

    private static ColumnType columnType(String type) {
        return switch (type) {
            case "note" -> ColumnType.COLUMN_TYPE_NOTE;
            case "text" -> ColumnType.COLUMN_TYPE_TEXT;
            case "choice" -> ColumnType.COLUMN_TYPE_CHOICE;
            case "multichoice" -> ColumnType.COLUMN_TYPE_MULTICHOICE;
            case "number" -> ColumnType.COLUMN_TYPE_NUMBER;
            case "currency" -> ColumnType.COLUMN_TYPE_CURRENCY;
            case "dateTime" -> ColumnType.COLUMN_TYPE_DATE_TIME;
            case "lookup" -> ColumnType.COLUMN_TYPE_LOOKUP;
            case "boolean" -> ColumnType.COLUMN_TYPE_BOOLEAN;
            case "user" -> ColumnType.COLUMN_TYPE_USER;
            case "url" -> ColumnType.COLUMN_TYPE_URL;
            case "calculated" -> ColumnType.COLUMN_TYPE_CALCULATED;
            case "location" -> ColumnType.COLUMN_TYPE_LOCATION;
            case "geolocation" -> ColumnType.COLUMN_TYPE_GEOLOCATION;
            case "term" -> ColumnType.COLUMN_TYPE_TERM;
            case "multiterm" -> ColumnType.COLUMN_TYPE_MULTITERM;
            case "thumbnail" -> ColumnType.COLUMN_TYPE_THUMBNAIL;
            case "approvalStatus" -> ColumnType.COLUMN_TYPE_APPROVAL_STATUS;
            case "" -> ColumnType.COLUMN_TYPE_UNSPECIFIED;
            default -> ColumnType.COLUMN_TYPE_UNKNOWN_FUTURE_VALUE;
        };
    }

    private static RetentionBehavior retentionBehavior(String value) {
        return switch (value) {
            case "doNotRetain" -> RetentionBehavior.RETENTION_BEHAVIOR_DO_NOT_RETAIN;
            case "retain" -> RetentionBehavior.RETENTION_BEHAVIOR_RETAIN;
            case "retainAsRecord" -> RetentionBehavior.RETENTION_BEHAVIOR_RETAIN_AS_RECORD;
            case "retainAsRegulatoryRecord" ->
                    RetentionBehavior.RETENTION_BEHAVIOR_RETAIN_AS_REGULATORY_RECORD;
            case "" -> RetentionBehavior.RETENTION_BEHAVIOR_UNSPECIFIED;
            default -> RetentionBehavior.RETENTION_BEHAVIOR_UNKNOWN_FUTURE_VALUE;
        };
    }

    // ---------------------------------------------------------------- helpers

    /**
     * RFC3339 / ISO-8601 to Timestamp. Tolerant of offset forms
     * ({@code +00:00} as well as {@code Z}); blank or unparseable input
     * yields null so the field stays unset.
     *
     * @param rfc3339 an RFC3339 / ISO-8601 timestamp, or blank
     * @return the proto timestamp, or {@code null} when input is blank or unparseable
     */
    public static Timestamp timestamp(String rfc3339) {
        if (rfc3339 == null || rfc3339.isBlank()) {
            return null;
        }
        String trimmed = rfc3339.trim();
        Instant instant;
        try {
            instant = Instant.parse(trimmed);
        } catch (DateTimeParseException e) {
            try {
                instant = OffsetDateTime.parse(trimmed).toInstant();
            } catch (DateTimeParseException e2) {
                return null;
            }
        }
        return Timestamp.newBuilder()
                .setSeconds(instant.getEpochSecond())
                .setNanos(instant.getNano())
                .build();
    }

    private static void ts(JsonNode node, String field, Consumer<Timestamp> setter) {
        Timestamp value = timestamp(text(node, field));
        if (value != null) {
            setter.accept(value);
        }
    }

    private static boolean truthy(JsonNode node) {
        if (node.isBoolean()) {
            return node.booleanValue();
        }
        return node.isTextual() && node.asText().equalsIgnoreCase("true");
    }

    private static long longValue(JsonNode node, String field) {
        JsonNode value = node.path(field);
        return value.isNumber() ? value.longValue() : 0L;
    }

    private static String firstText(JsonNode node, String... fields) {
        for (String field : fields) {
            String value = text(node, field);
            if (!value.isEmpty()) {
                return value;
            }
        }
        return "";
    }

    private static String text(JsonNode node, String field) {
        JsonNode value = node.path(field);
        return value.isTextual() ? value.asText() : "";
    }
}
