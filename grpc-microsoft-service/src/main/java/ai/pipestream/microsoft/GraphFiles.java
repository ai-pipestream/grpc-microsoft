package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

/**
 * Files and their metadata over Microsoft Graph — OneDrive and SharePoint Online through
 * the same {@code driveItem} model (a OneDrive for Business drive <em>is</em> a SharePoint
 * document library under the hood, so everything here works on whichever the tenant has).
 * {@link #listItemFields} returns the SharePoint column values of a document; the mapper
 * flattens them into typed {@code ListColumn}s. The write side uploads content (simple
 * PUT or an upload session above {@link #SIMPLE_UPLOAD_LIMIT}) and can patch list columns.
 */
public final class GraphFiles {

    /** Simple-upload ceiling; larger files use {@link #uploadSession}. */
    public static final int SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;

    private final GraphClient graph;

    /**
     * Binds this files API to an authorized Graph client.
     *
     * @param graph the authorized REST client
     */
    public GraphFiles(GraphClient graph) {
        this.graph = Objects.requireNonNull(graph, "graph");
    }

    /**
     * The signed-in user (delegated flows) — the cheapest connectivity probe.
     *
     * @return the {@code /me} JSON object
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode me() throws IOException, InterruptedException {
        return graph.get("/me");
    }

    /**
     * The signed-in user's OneDrive (its {@code driveType} says personal vs
     * business).
     *
     * @return the {@code /me/drive} JSON object
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode meDrive() throws IOException, InterruptedException {
        return graph.get("/me/drive");
    }

    /**
     * SharePoint site search; an empty result on a OneDrive-only tenant, not an
     * error.
     *
     * @param query search string
     * @return the sites search JSON page
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode searchSites(String query) throws IOException, InterruptedException {
        return graph.get("/sites?search=" + URLEncoder.encode(query, StandardCharsets.UTF_8));
    }

    /**
     * Document libraries (drives) of a SharePoint site.
     *
     * @param siteId SharePoint site id
     * @return the drives JSON page
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode drives(String siteId) throws IOException, InterruptedException {
        return graph.get("/sites/" + siteId + "/drives");
    }

    /**
     * One drive by id.
     *
     * @param driveId drive id
     * @return the drive JSON object
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode drive(String driveId) throws IOException, InterruptedException {
        return graph.get("/drives/" + driveId);
    }

    /**
     * One drive item by id.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @return the driveItem JSON object
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode item(String driveId, String itemId)
            throws IOException, InterruptedException {
        return graph.get("/drives/" + driveId + "/items/" + itemId);
    }

    /**
     * Children of a folder; {@code folderPath} null or "/" lists the drive root.
     *
     * @param driveId parent drive id
     * @param folderPath folder path within the drive; {@code null} or {@code "/"}
     *        is the root
     * @return one page of child driveItems
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode children(String driveId, String folderPath)
            throws IOException, InterruptedException {
        return graph.get(childrenPath(driveId, folderPath));
    }

    /**
     * Follows {@code @odata.nextLink} until the folder listing is exhausted.
     *
     * @param driveId parent drive id
     * @param folderPath folder path within the drive; {@code null} or {@code "/"}
     *        is the root
     * @return every child driveItem across pages
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public java.util.List<JsonNode> childrenAll(String driveId, String folderPath)
            throws IOException, InterruptedException {
        java.util.List<JsonNode> items = new java.util.ArrayList<>();
        JsonNode page = children(driveId, folderPath);
        while (true) {
            page.path("value").forEach(items::add);
            String next = page.path("@odata.nextLink").asText("");
            if (next.isBlank()) {
                return items;
            }
            page = graph.get(next);
        }
    }

    static String childrenPath(String driveId, String folderPath) {
        String base = "/drives/" + driveId + "/root";
        return folderPath == null || folderPath.isBlank() || folderPath.equals("/")
                ? base + "/children"
                : base + ":" + encodePath(folderPath) + ":/children";
    }

    /**
     * Downloads file content for one drive item.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @return the file bytes
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public byte[] download(String driveId, String itemId)
            throws IOException, InterruptedException {
        return graph.getBytes("/drives/" + driveId + "/items/" + itemId + "/content");
    }

    /**
     * Uploads (or overwrites) a file by path. Content up to {@link #SIMPLE_UPLOAD_LIMIT};
     * the destination is always the caller's explicit drive and path.
     *
     * @param driveId destination drive id
     * @param folderPath destination folder; {@code null} or {@code "/"} is the root
     * @param fileName file name within the folder
     * @param content file bytes
     * @param contentType MIME type; {@code null} becomes
     *        {@code application/octet-stream}
     * @return the created or updated driveItem JSON
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode upload(String driveId, String folderPath, String fileName, byte[] content,
                           String contentType) throws IOException, InterruptedException {
        if (content.length > SIMPLE_UPLOAD_LIMIT) {
            throw new IllegalArgumentException("Content is " + content.length + " bytes; the "
                    + "simple upload lane caps at " + SIMPLE_UPLOAD_LIMIT
                    + " - use an upload session for large files");
        }
        String path = (folderPath == null || folderPath.isBlank() || folderPath.equals("/")
                ? "" : normalize(folderPath)) + "/" + fileName;
        return graph.putBytes("/drives/" + driveId + "/root:" + encodePath(path) + ":/content",
                content, contentType == null ? "application/octet-stream" : contentType);
    }

    /**
     * Uploads by path, using a simple PUT under {@link #SIMPLE_UPLOAD_LIMIT}
     * and an upload session for larger files. Destination folders in the path
     * are created as needed by Graph.
     *
     * @param driveId destination drive id
     * @param destPath file path from drive root, for example {@code /okf/pages/200.md}
     * @param content file bytes
     * @param contentType MIME type; {@code null} becomes {@code application/octet-stream}
     * @return the created or updated driveItem JSON (last upload-session response)
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode uploadOrSession(String driveId, String destPath, byte[] content,
            String contentType) throws IOException, InterruptedException {
        String mime = contentType == null ? "application/octet-stream" : contentType;
        if (content.length <= SIMPLE_UPLOAD_LIMIT) {
            String path = normalize(destPath);
            return graph.putBytes("/drives/" + driveId + "/root:" + encodePath(path) + ":/content",
                    content, mime);
        }
        return uploadSession(driveId, destPath, content, mime);
    }

    /**
     * Upload-session PUT for files larger than {@link #SIMPLE_UPLOAD_LIMIT}.
     * Fragments are a multiple of 320 KiB except the last. The upload URL is
     * called without {@code Authorization}.
     *
     * @param driveId destination drive id
     * @param destPath file path from drive root
     * @param content file bytes
     * @param contentType MIME type
     * @return the completed driveItem JSON
     * @throws IOException if Graph returns an error or omits {@code uploadUrl}
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode uploadSession(String driveId, String destPath, byte[] content,
            String contentType) throws IOException, InterruptedException {
        String path = normalize(destPath);
        ObjectNode body = GraphClient.object();
        body.putObject("item")
                .put("@microsoft.graph.conflictBehavior", "replace")
                .put("name", fileName(path));
        JsonNode session = graph.post(
                "/drives/" + driveId + "/root:" + encodePath(path) + ":/createUploadSession",
                body);
        String uploadUrl = session.path("uploadUrl").asText("");
        if (uploadUrl.isBlank()) {
            throw new IOException("createUploadSession returned no uploadUrl for " + destPath);
        }
        final int fragment = 320 * 1024 * 10;
        int offset = 0;
        JsonNode last = GraphClient.object();
        while (offset < content.length) {
            int end = Math.min(offset + fragment, content.length);
            byte[] chunk = java.util.Arrays.copyOfRange(content, offset, end);
            String range = "bytes " + offset + "-" + (end - 1) + "/" + content.length;
            last = graph.putRangeUnauthenticated(uploadUrl, chunk, contentType, range);
            offset = end;
        }
        return last;
    }

    /**
     * The SharePoint list-item column values behind a document — titles, choice columns,
     * managed metadata, whatever the library declares. This is the metadata read lane.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @return the listItem JSON including expanded {@code fields}
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode listItemFields(String driveId, String itemId)
            throws IOException, InterruptedException {
        return graph.get("/drives/" + driveId + "/items/" + itemId
                + "/listItem?$expand=fields");
    }

    /**
     * Just the list-item columns behind a document — the {@code fields} object out of
     * {@link #listItemFields}. Returns an empty object when the item has no list item (a personal-OneDrive file that
     * belongs to no document library), so a caller can sample a folder without null checks.
     * A {@code driveId} or {@code itemId} that does not resolve still fails.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @return the {@code fields} object, or empty when the item has no list item
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public ObjectNode listItemFieldsOnly(String driveId, String itemId)
            throws IOException, InterruptedException {
        JsonNode fields;
        try {
            fields = listItemFields(driveId, itemId).path("fields");
        } catch (GraphClient.GraphApiException e) {
            // Graph answers 404 both for a file with no backing list item (a plain
            // personal-OneDrive file: no columns) and for an item that is not there at all;
            // only the item itself resolving tells the two apart.
            if (e.status() == 404 && itemExists(driveId, itemId)) {
                return JsonNodeFactory.instance.objectNode();
            }
            throw e;
        }
        return fields.isObject() ? (ObjectNode) fields : JsonNodeFactory.instance.objectNode();
    }

    /**
     * List-item columns, or empty on 404 (a personal-OneDrive file, or a
     * library item with no list item). Use this from a crawler that already
     * holds the drive item; it does not re-GET the item to distinguish 404
     * causes.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @return the {@code fields} object, or empty
     * @throws IOException if Graph returns a non-404 error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public ObjectNode listItemFieldsOrEmpty(String driveId, String itemId)
            throws IOException, InterruptedException {
        try {
            JsonNode fields = listItemFields(driveId, itemId).path("fields");
            return fields.isObject() ? (ObjectNode) fields : JsonNodeFactory.instance.objectNode();
        } catch (GraphClient.GraphApiException e) {
            if (e.status() == 404) {
                return JsonNodeFactory.instance.objectNode();
            }
            throw e;
        }
    }

    /** Whether the driveItem resolves; any error resolving it leaves the caller's own to report. */
    private boolean itemExists(String driveId, String itemId)
            throws IOException, InterruptedException {
        try {
            graph.get("/drives/" + driveId + "/items/" + itemId);
            return true;
        } catch (GraphClient.GraphApiException e) {
            return false;
        }
    }

    /**
     * Patches list-item columns; {@code fields} holds exactly the columns to change.
     *
     * @param driveId parent drive id
     * @param itemId drive-item id
     * @param fields columns to patch
     * @return the patched fields JSON
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode updateListItemFields(String driveId, String itemId, ObjectNode fields)
            throws IOException, InterruptedException {
        return graph.patch("/drives/" + driveId + "/items/" + itemId + "/listItem/fields",
                fields);
    }

    private static String normalize(String folderPath) {
        String cleaned = folderPath.replace('\\', '/');
        if (!cleaned.startsWith("/")) {
            cleaned = "/" + cleaned;
        }
        return cleaned.replaceAll("/+$", "");
    }

    private static String fileName(String destPath) {
        String normalized = normalize(destPath);
        int slash = normalized.lastIndexOf('/');
        return slash < 0 ? normalized : normalized.substring(slash + 1);
    }

    private static String encodePath(String path) {
        StringBuilder encoded = new StringBuilder();
        for (String segment : normalize(path).split("/")) {
            if (segment.isEmpty()) {
                continue;
            }
            encoded.append('/').append(URLEncoder.encode(segment, StandardCharsets.UTF_8)
                    .replace("+", "%20"));
        }
        return encoded.toString();
    }
}
