package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.Site;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.Timestamp;

import java.time.Instant;
import java.time.OffsetDateTime;
import java.time.format.DateTimeParseException;

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
                .setFolder(node.path("folder").isObject());
        JsonNode file = node.path("file");
        if (file.isObject()) {
            b.setMimeType(text(file, "mimeType"));
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
