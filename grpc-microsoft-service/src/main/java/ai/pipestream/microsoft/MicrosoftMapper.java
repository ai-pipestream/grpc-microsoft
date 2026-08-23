package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.FileHashes;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.Site;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.Timestamp;

import java.time.Instant;
import java.time.OffsetDateTime;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Jackson {@link JsonNode} to domain proto translation for Microsoft Graph
 * drive, site, user, and driveItem shapes. Missing fields stay unset; unknown
 * values never throw.
 */
public final class MicrosoftMapper {

    /** Creates a mapper. */
    public MicrosoftMapper() {
    }

    /**
     * Maps a Graph site JSON object to the domain proto.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public Site toSite(JsonNode node) {
        return Site.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDisplayName(text(node, "displayName"))
                .setWebUrl(text(node, "webUrl"))
                .build();
    }

    /**
     * Maps a Graph drive JSON object to the domain proto.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public Drive toDrive(JsonNode node) {
        return Drive.newBuilder()
                .setId(text(node, "id"))
                .setName(text(node, "name"))
                .setDriveType(text(node, "driveType"))
                .setWebUrl(text(node, "webUrl"))
                .build();
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

    /**
     * Maps a Graph user JSON object to the domain proto.
     *
     * @param node the API object
     * @return the proto; missing fields stay unset
     */
    public GraphUser toUser(JsonNode node) {
        return GraphUser.newBuilder()
                .setId(text(node, "id"))
                .setDisplayName(text(node, "displayName"))
                .setUserPrincipalName(text(node, "userPrincipalName"))
                .setMail(text(node, "mail"))
                .build();
    }

    /**
     * Maps a Graph driveItem JSON object to the domain proto.
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
                .setSize(node.path("size").isIntegralNumber() ? node.path("size").longValue() : 0L)
                .setFolder(node.path("folder").isObject())
                .setDescription(text(node, "description"))
                .setEtag(firstText(node, "eTag", "etag"));
        JsonNode folder = node.path("folder");
        if (folder.isObject() && folder.path("childCount").isIntegralNumber()) {
            b.setChildCount(folder.path("childCount").intValue());
        }
        JsonNode file = node.path("file");
        if (file.isObject()) {
            b.setMimeType(text(file, "mimeType"));
            JsonNode hashes = file.path("hashes");
            if (hashes.isObject()) {
                b.setHashes(FileHashes.newBuilder()
                        .setSha1(firstText(hashes, "sha1Hash", "sha1"))
                        .setSha256(firstText(hashes, "sha256Hash", "sha256"))
                        .setQuickXor(firstText(hashes, "quickXorHash", "quickXor"))
                        .setCrc32(firstText(hashes, "crc32Hash", "crc32"))
                        .build());
            }
        }
        JsonNode parent = node.path("parentReference");
        if (parent.isObject()) {
            b.setParentId(text(parent, "id"));
            if (b.getDriveId().isEmpty()) {
                b.setDriveId(text(parent, "driveId"));
            }
        }
        Timestamp created = timestamp(text(node, "createdDateTime"));
        if (created != null) {
            b.setCreatedAt(created);
        }
        Timestamp modified = timestamp(text(node, "lastModifiedDateTime"));
        if (modified != null) {
            b.setLastModifiedAt(modified);
        }
        if (node.path("createdBy").isObject()) {
            b.setCreatedBy(toIdentitySet(node.path("createdBy")));
        }
        if (node.path("lastModifiedBy").isObject()) {
            b.setLastModifiedBy(toIdentitySet(node.path("lastModifiedBy")));
        }
        return b.build();
    }

    /**
     * Copies {@code item} and replaces {@code list_columns} with the flattened
     * SharePoint fields object. {@code @odata.*} keys are skipped; nested
     * objects flatten one level as {@code Parent.Child}.
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
     * Flattens a SharePoint {@code fields} object into typed {@link ListColumn}s.
     *
     * @param fields Graph fields object; {@code null} or non-object yields empty
     * @return columns in field-name order
     */
    public List<ListColumn> toListColumns(JsonNode fields) {
        List<ListColumn> columns = new ArrayList<>();
        if (fields == null || !fields.isObject()) {
            return columns;
        }
        Iterator<Map.Entry<String, JsonNode>> fieldsIt = fields.fields();
        while (fieldsIt.hasNext()) {
            Map.Entry<String, JsonNode> entry = fieldsIt.next();
            addColumns(columns, entry.getKey(), entry.getValue(), true);
        }
        return columns;
    }

    private static void addColumns(List<ListColumn> out, String name, JsonNode value,
            boolean flattenObjects) {
        if (name == null || name.startsWith("@odata.") || name.startsWith("@")) {
            return;
        }
        if (value == null || value.isNull() || value.isMissingNode()) {
            return;
        }
        if (value.isObject() && flattenObjects) {
            Iterator<Map.Entry<String, JsonNode>> nested = value.fields();
            while (nested.hasNext()) {
                Map.Entry<String, JsonNode> entry = nested.next();
                addColumns(out, name + "." + entry.getKey(), entry.getValue(), false);
            }
            return;
        }
        ListColumn column = toColumn(name, value);
        if (column != null) {
            out.add(column);
        }
    }

    private static ListColumn toColumn(String name, JsonNode value) {
        ListColumn.Builder b = ListColumn.newBuilder().setName(name);
        if (value.isBoolean()) {
            return b.setBoolValue(value.booleanValue()).build();
        }
        if (value.isIntegralNumber()) {
            return b.setIntValue(value.longValue()).build();
        }
        if (value.isFloatingPointNumber()) {
            return b.setDoubleValue(value.doubleValue()).build();
        }
        if (value.isArray()) {
            List<String> parts = new ArrayList<>();
            for (JsonNode element : value) {
                if (element.isValueNode() && !element.isNull()) {
                    parts.add(element.asText());
                }
            }
            if (parts.isEmpty()) {
                return null;
            }
            return b.setStringValue(String.join(", ", parts)).build();
        }
        if (value.isTextual()) {
            String text = value.asText();
            Timestamp ts = timestamp(text);
            if (ts != null && looksLikeTimestamp(text)) {
                return b.setTimestampValue(ts).build();
            }
            return b.setStringValue(text).build();
        }
        if (value.isObject()) {
            return null;
        }
        return b.setStringValue(value.asText()).build();
    }

    private static boolean looksLikeTimestamp(String text) {
        return text.length() >= 10 && (text.charAt(4) == '-' || text.endsWith("Z")
                || text.contains("T"));
    }

    private static IdentitySet toIdentitySet(JsonNode node) {
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
        return b.build();
    }

    private static Identity toIdentity(JsonNode node) {
        return Identity.newBuilder()
                .setId(text(node, "id"))
                .setDisplayName(text(node, "displayName"))
                .build();
    }

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
