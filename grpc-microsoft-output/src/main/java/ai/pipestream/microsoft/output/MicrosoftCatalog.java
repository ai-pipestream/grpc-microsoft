package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.SharePointIdentitySet;
import ai.pipestream.microsoft.v1.SharePointList;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.OkfActor;
import ai.pipestream.okf.OkfConcept;
import ai.pipestream.okf.OkfPaths;
import ai.pipestream.okf.OkfYaml;
import ai.pipestream.okf.warc.WarcDigest;
import com.google.protobuf.Timestamp;

import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.util.List;
import java.util.Optional;

/**
 * Maps Microsoft Graph protobuf entities onto OKF catalog entries. Live
 * {@code web_url} / {@code download_url} is {@code WARC-Target-URI}. SharePoint
 * list columns are a markdown table in the concept body, not a JSON blob.
 */
public final class MicrosoftCatalog {

    /** Actor recorded on generated concepts. */
    public static final String ACTOR = OkfActor.process("grpc-microsoft/okf-producer");

    private MicrosoftCatalog() {
    }

    /**
     * Maps a change. Deletes become {@code status: deprecated}.
     *
     * @param change a crawler change
     * @return the catalog entry, or empty when the change has no entity
     */
    public static Optional<CatalogEntry> from(MicrosoftChange change) {
        if (!change.hasEntity()) {
            return Optional.empty();
        }
        CatalogEntry entry = from(change.getEntity());
        if (change.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE) {
            entry = new CatalogEntry(entry.path(),
                    withStatus(entry.concept(), OkfConcept.Status.DEPRECATED),
                    entry.targetUri(), entry.mediaType(), entry.resourceBody(), entry.kind());
        }
        return Optional.of(entry);
    }

    /**
     * Maps an entity.
     *
     * @param entity wrapped Graph entity
     * @return the catalog entry
     */
    public static CatalogEntry from(MicrosoftEntity entity) {
        Instant at = instant(entity.getIngestedAt());
        return switch (entity.getEntityCase()) {
            case SITE -> site(entity.getSite(), at);
            case DRIVE -> drive(entity.getDrive(), at);
            case DRIVE_ITEM -> driveItem(entity.getDriveItem(), at);
            case USER -> user(entity.getUser(), at);
            case LIST -> list(entity.getList(), at);
            case LIST_ITEM -> listItem(entity.getListItem(), at);
            case PERMISSION -> permission(entity.getEntityId(), entity.getPermission(), at);
            case ENTITY_NOT_SET -> fallback(entity, at);
        };
    }

    private static CatalogEntry site(Site site, Instant at) {
        String path = OkfPaths.join("sites", OkfPaths.segment(site.getId()) + ".md");
        String uri = uriOrUrn(site.getWebUrl(), "site", site.getId());
        String title = site.getDisplayName().isBlank() ? site.getName() : site.getDisplayName();
        OkfConcept concept = base("Site", title, site.getName(), uri, at)
                .extra("graph_id", site.getId())
                .body("SharePoint site `" + site.getName() + "`.")
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", title, "site");
    }

    private static CatalogEntry drive(Drive drive, Instant at) {
        String path = OkfPaths.join("drives", OkfPaths.segment(drive.getId()) + ".md");
        String uri = uriOrUrn(drive.getWebUrl(), "drive", drive.getId());
        OkfConcept concept = base("Drive", drive.getName(), drive.getDriveType(), uri, at)
                .extra("graph_id", drive.getId())
                .extra("drive_type", drive.getDriveType())
                .extra("site_id", drive.getSiteId())
                .body("Drive type `" + drive.getDriveType() + "`.")
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", drive.getName(), "drive");
    }

    private static CatalogEntry user(GraphUser user, Instant at) {
        String path = OkfPaths.join("users", OkfPaths.segment(user.getId()) + ".md");
        String uri = uriOrUrn("", "user", user.getId());
        String title = user.getDisplayName().isBlank() ? user.getId() : user.getDisplayName();
        OkfConcept concept = base("User", title, user.getMail(), uri, at)
                .extra("graph_id", user.getId())
                .extra("user_principal_name", user.getUserPrincipalName())
                .extra("mail", user.getMail())
                .body(title)
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", title, "user");
    }

    private static CatalogEntry driveItem(DriveItem item, Instant at) {
        String path = OkfPaths.join("items", OkfPaths.segment(item.getDriveId()),
                OkfPaths.segment(item.getId()) + ".md");
        String uri = firstNonBlank(item.getDownloadUrl(), item.getWebUrl());
        uri = uriOrUrn(uri, "item", item.getDriveId() + "/" + item.getId());
        byte[] payload;
        String media;
        if (item.hasContent() && !item.getContent().isEmpty()) {
            payload = item.getContent().toByteArray();
            media = item.getMimeType().isBlank() ? "application/octet-stream" : item.getMimeType();
        } else {
            payload = item.getName().getBytes(StandardCharsets.UTF_8);
            media = "text/plain";
        }
        String kind = item.getFolder() ? "folder" : "file";
        StringBuilder body = new StringBuilder();
        if (!item.getDescription().isBlank()) {
            body.append(item.getDescription()).append("\n\n");
        }
        body.append("* Kind: ").append(kind).append('\n');
        body.append("* Size: ").append(item.getSize()).append('\n');
        if (!item.getMimeType().isBlank()) {
            body.append("* MIME: ").append(item.getMimeType()).append('\n');
        }
        if (!item.getEtag().isBlank()) {
            body.append("* eTag: `").append(item.getEtag()).append("`\n");
        }
        if (item.getFolder()) {
            body.append("* Child count: ").append(item.getChildCount()).append('\n');
        }
        appendIdentity(body, "Created by", item.getCreatedBy());
        appendIdentity(body, "Last modified by", item.getLastModifiedBy());
        if (item.hasHashes()) {
            if (!item.getHashes().getSha1().isBlank()) {
                body.append("* SHA-1: `").append(item.getHashes().getSha1()).append("`\n");
            }
            if (!item.getHashes().getSha256().isBlank()) {
                body.append("* SHA-256: `").append(item.getHashes().getSha256()).append("`\n");
            }
            if (!item.getHashes().getQuickXor().isBlank()) {
                body.append("* QuickXOR: `").append(item.getHashes().getQuickXor()).append("`\n");
            }
            if (!item.getHashes().getCrc32().isBlank()) {
                body.append("* CRC32: `").append(item.getHashes().getCrc32()).append("`\n");
            }
        }
        if (item.getListColumnsCount() > 0) {
            body.append("\n## SharePoint columns\n\n");
            body.append("| Column | Value |\n| --- | --- |\n");
            for (ListColumn column : item.getListColumnsList()) {
                body.append("| ").append(escapeCell(column.getName())).append(" | ")
                        .append(escapeCell(columnValue(column))).append(" |\n");
            }
        }
        OkfConcept.Builder b = base(item.getFolder() ? "Folder" : "File", item.getName(),
                item.getDescription().isBlank() ? item.getMimeType() : item.getDescription(),
                uri, at)
                .tag(kind)
                .extra("graph_id", item.getId())
                .extra("drive_id", item.getDriveId())
                .extra("parent_id", item.getParentId())
                .extra("etag", item.getEtag())
                .extra("mime_type", item.getMimeType())
                .extra("size", Long.toString(item.getSize()))
                .body(body.toString());
        if (item.hasContent() && !item.getContent().isEmpty()) {
            b.extra("sha256", WarcDigest.sha256Hex(payload));
        }
        return new CatalogEntry(path, b.build(), uri, media, payload, kind);
    }

    private static CatalogEntry list(SharePointList list, Instant at) {
        String path = OkfPaths.join("lists", OkfPaths.segment(list.getId()) + ".md");
        String uri = uriOrUrn(list.getWebUrl(), "list", list.getId());
        String title = list.getDisplayName().isBlank() ? list.getName() : list.getDisplayName();
        StringBuilder body = new StringBuilder();
        if (!list.getDescription().isBlank()) {
            body.append(list.getDescription()).append("\n\n");
        }
        body.append("* Template: ").append(list.getList().getTemplate()).append('\n');
        body.append("* Hidden: ").append(list.getList().getHidden()).append('\n');
        appendIdentity(body, "Created by", list.getCreatedBy());
        OkfConcept concept = base("List", title, list.getDescription(), uri, at)
                .extra("graph_id", list.getId())
                .extra("site_id", list.getSiteId())
                .extra("template", list.getList().getTemplate())
                .extra("drive_id", list.getDriveId())
                .body(body.toString())
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", title, "list");
    }

    private static CatalogEntry listItem(ListItem item, Instant at) {
        String path = OkfPaths.join("lists", OkfPaths.segment(item.getListId()), "items",
                OkfPaths.segment(item.getId()) + ".md");
        String uri = uriOrUrn(item.getWebUrl(), "list_item", item.getListId() + "/" + item.getId());
        String title = firstNonBlank(item.getName(), titleColumn(item));
        if (title.isBlank()) {
            title = item.getId();
        }
        StringBuilder body = new StringBuilder();
        if (!item.getContentType().getName().isBlank()) {
            body.append("* Content type: ").append(item.getContentType().getName()).append('\n');
        }
        appendIdentity(body, "Created by", item.getCreatedBy());
        appendIdentity(body, "Last modified by", item.getLastModifiedBy());
        appendColumns(body, item.getFieldsList());
        OkfConcept concept = base("ListItem", title, item.getContentType().getName(), uri, at)
                .extra("graph_id", item.getId())
                .extra("site_id", item.getSiteId())
                .extra("list_id", item.getListId())
                .extra("etag", item.getEtag())
                .body(body.toString())
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", title, "list_item");
    }

    private static CatalogEntry permission(String entityId, Permission permission, Instant at) {
        String path = OkfPaths.join("permissions", OkfPaths.segment(entityId) + ".md");
        String uri = uriOrUrn("", "permission", entityId);
        String roles = String.join(", ", permission.getRolesList());
        StringBuilder body = new StringBuilder();
        body.append("* Roles: ").append(roles).append('\n');
        if (permission.hasLink()) {
            body.append("* Link: ").append(permission.getLink().getType()).append(" / ")
                    .append(permission.getLink().getScope()).append('\n');
        }
        if (permission.hasGrantedTo()) {
            appendGrantee(body, permission.getGrantedTo());
        }
        for (SharePointIdentitySet grantee : permission.getGrantedToIdentitiesList()) {
            appendGrantee(body, grantee);
        }
        body.append("* Inherited: ").append(permission.hasInheritedFrom()).append('\n');
        OkfConcept concept = base("Permission", roles.isBlank() ? permission.getId() : roles, roles,
                uri, at)
                .extra("graph_id", permission.getId())
                .body(body.toString())
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", permission.getId(), "permission");
    }

    private static void appendGrantee(StringBuilder body, SharePointIdentitySet grantee) {
        for (Identity identity : List.of(grantee.getUser(), grantee.getGroup(),
                grantee.getApplication(), grantee.getDevice())) {
            if (!identity.getId().isBlank() || !identity.getDisplayName().isBlank()) {
                body.append("* Grantee: ").append(identity.getDisplayName())
                        .append(" (`").append(identity.getId()).append("`)\n");
            }
        }
        if (grantee.hasSharePointGroup()) {
            body.append("* Grantee: SharePoint group ").append(grantee.getSharePointGroup().getDisplayName())
                    .append(" (`").append(grantee.getSharePointGroup().getId()).append("`)\n");
        }
        if (grantee.hasSiteGroup()) {
            body.append("* Grantee: site group ").append(grantee.getSiteGroup().getDisplayName())
                    .append(" (`").append(grantee.getSiteGroup().getId()).append("`)\n");
        }
    }

    private static String titleColumn(ListItem item) {
        for (ListColumn column : item.getFieldsList()) {
            if (column.getName().equals("Title")) {
                return columnValue(column);
            }
        }
        return "";
    }

    private static void appendColumns(StringBuilder body, List<ListColumn> columns) {
        if (columns.isEmpty()) {
            return;
        }
        body.append("\n## SharePoint columns\n\n");
        body.append("| Column | Value |\n| --- | --- |\n");
        for (ListColumn column : columns) {
            body.append("| ").append(escapeCell(column.getName())).append(" | ")
                    .append(escapeCell(columnValue(column))).append(" |\n");
        }
    }

    private static CatalogEntry fallback(MicrosoftEntity entity, Instant at) {
        String id = entity.getEntityId();
        String path = OkfPaths.join("entities", OkfPaths.segment(id) + ".md");
        String uri = uriOrUrn("", "entity", id);
        OkfConcept concept = base("Entity", id, "", uri, at)
                .extra("graph_id", id)
                .body("Captured Graph entity `" + id + "`.")
                .build();
        return CatalogEntry.text(path, concept, uri, "text/plain", id, "entity");
    }

    private static OkfConcept.Builder base(String type, String title, String description,
            String resource, Instant at) {
        return OkfConcept.of(type)
                .title(title)
                .description(description)
                .resource(resource)
                .generated(new OkfConcept.Generated(ACTOR, at))
                .source(new OkfConcept.Source(null, resource, title, ACTOR, null, at, null));
    }

    private static OkfConcept withStatus(OkfConcept concept, OkfConcept.Status status) {
        OkfConcept.Builder b = OkfConcept.of(concept.type())
                .title(concept.title().orElse(null))
                .description(concept.description().orElse(null))
                .resource(concept.resource().orElse(null))
                .tags(concept.tags())
                .status(status)
                .body(concept.body());
        concept.generated().ifPresent(b::generated);
        concept.sources().forEach(b::source);
        concept.extra().forEach(b::extra);
        return b.build();
    }

    private static void appendIdentity(StringBuilder body, String label, IdentitySet set) {
        if (set == null) {
            return;
        }
        Identity user = set.getUser();
        if (user.getId().isEmpty() && user.getDisplayName().isEmpty()) {
            return;
        }
        body.append("* ").append(label).append(": ");
        if (!user.getDisplayName().isBlank()) {
            body.append(user.getDisplayName());
        }
        if (!user.getId().isBlank()) {
            body.append(" (`").append(user.getId()).append("`)");
        }
        body.append('\n');
    }

    private static String columnValue(ListColumn column) {
        return switch (column.getValueCase()) {
            case STRING_VALUE -> column.getStringValue();
            case INT_VALUE -> Long.toString(column.getIntValue());
            case DOUBLE_VALUE -> Double.toString(column.getDoubleValue());
            case BOOL_VALUE -> Boolean.toString(column.getBoolValue());
            case TIMESTAMP_VALUE -> OkfYaml.iso(instant(column.getTimestampValue()));
            case VALUE_NOT_SET -> "";
        };
    }

    private static String escapeCell(String value) {
        if (value == null) {
            return "";
        }
        return value.replace("|", "\\|").replace("\n", " ");
    }

    private static String uriOrUrn(String url, String kind, String id) {
        if (url != null && !url.isBlank()) {
            return url;
        }
        return "urn:okf:0.2:microsoft:" + kind + ":" + (id == null || id.isBlank() ? "_" : id);
    }

    private static String firstNonBlank(String a, String b) {
        if (a != null && !a.isBlank()) {
            return a;
        }
        return b == null ? "" : b;
    }

    private static Instant instant(Timestamp timestamp) {
        if (timestamp == null || (timestamp.getSeconds() == 0 && timestamp.getNanos() == 0)) {
            return Instant.now();
        }
        return Instant.ofEpochSecond(timestamp.getSeconds(), timestamp.getNanos());
    }
}
