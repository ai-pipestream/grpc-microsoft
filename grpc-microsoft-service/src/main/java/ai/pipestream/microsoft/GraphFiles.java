package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Map;
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

    /** Bytes one upload-session request must be a multiple of, except the last. */
    public static final int FRAGMENT_UNIT_BYTES = 320 * 1024;

    /** Upload-session request size used when the caller names none. */
    public static final int DEFAULT_FRAGMENT_BYTES = 32 * FRAGMENT_UNIT_BYTES;

    /** Largest upload-session request Graph accepts. */
    public static final int MAX_FRAGMENT_BYTES = 60 * 1024 * 1024;

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
     * The authorized REST client this API wraps.
     *
     * @return the client
     */
    public GraphClient client() {
        return graph;
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
        java.util.Set<String> seen = new java.util.HashSet<>();
        while (true) {
            page.path("value").forEach(items::add);
            String next = page.path("@odata.nextLink").asText("");
            if (next.isBlank()) {
                return items;
            }
            page = graph.get(GraphPager.admit(seen, next), Map.of());
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
        String uploadUrl = createUploadSession(
                "/drives/" + driveId + "/root:" + encodePath(path) + ":", fileName(path),
                "replace", "");
        final int fragment = DEFAULT_FRAGMENT_BYTES;
        int offset = 0;
        JsonNode last = GraphClient.object();
        while (offset < content.length) {
            int end = Math.min(offset + fragment, content.length);
            byte[] chunk = java.util.Arrays.copyOfRange(content, offset, end);
            last = putFragment(uploadUrl, chunk, chunk.length, offset, content.length);
            offset = end;
        }
        return last;
    }

    /**
     * Stores a file with one PUT.
     *
     * @param itemRoute the file's address ({@link GraphPaths#childByName})
     * @param content file bytes
     * @param length how many bytes of {@code content} to send
     * @param contentType MIME type; blank becomes {@code application/octet-stream}
     * @param conflictBehavior Graph {@code conflictBehavior}; blank for Graph's default
     * @param ifMatch eTag the existing file must have; blank for unconditional
     * @return the created or updated driveItem JSON
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putContent(String itemRoute, byte[] content, int length, String contentType,
            String conflictBehavior, String ifMatch) throws IOException, InterruptedException {
        String route = itemRoute + "/content";
        if (conflictBehavior != null && !conflictBehavior.isBlank()) {
            route += "?@microsoft.graph.conflictBehavior=" + conflictBehavior;
        }
        return graph.putBytes(route, content, length,
                contentType == null || contentType.isBlank() ? "application/octet-stream"
                        : contentType,
                ifMatch == null || ifMatch.isBlank() ? Map.of() : Map.of("if-match", ifMatch));
    }

    /**
     * Opens an upload session for a file.
     *
     * @param itemRoute the file's address ({@link GraphPaths#childByName})
     * @param name file name
     * @param conflictBehavior Graph {@code conflictBehavior}; blank for Graph's default
     * @param ifMatch eTag the existing file must have; blank for unconditional
     * @return the session's upload URL
     * @throws IOException if Graph returns an error or omits {@code uploadUrl}
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public String createUploadSession(String itemRoute, String name, String conflictBehavior,
            String ifMatch) throws IOException, InterruptedException {
        ObjectNode body = GraphClient.object();
        ObjectNode item = body.putObject("item");
        if (conflictBehavior != null && !conflictBehavior.isBlank()) {
            item.put("@microsoft.graph.conflictBehavior", conflictBehavior);
        }
        item.put("name", name);
        JsonNode session = graph.execute(new GraphClient.Call("POST",
                itemRoute + "/createUploadSession",
                ifMatch == null || ifMatch.isBlank() ? Map.of() : Map.of("if-match", ifMatch),
                body)).body();
        String uploadUrl = session.path("uploadUrl").asText("");
        if (uploadUrl.isBlank()) {
            throw new IOException("createUploadSession returned no uploadUrl for " + name);
        }
        return uploadUrl;
    }

    /**
     * Sends one fragment of an upload session.
     *
     * @param uploadUrl the session's upload URL
     * @param content fragment bytes
     * @param length how many bytes of {@code content} to send
     * @param offset position of the fragment in the file
     * @param total size of the whole file
     * @return the session status JSON, or the driveItem JSON after the last fragment
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public JsonNode putFragment(String uploadUrl, byte[] content, int length, long offset,
            long total) throws IOException, InterruptedException {
        String range = "bytes " + offset + "-" + (offset + length - 1) + "/" + total;
        return graph.putRangeUnauthenticated(uploadUrl, content, length,
                "application/octet-stream", range);
    }

    /**
     * Reads where an upload session expects the next byte.
     *
     * @param uploadUrl the session's upload URL
     * @return the next expected offset
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public long uploadSessionOffset(String uploadUrl) throws IOException, InterruptedException {
        return nextExpectedOffset(graph.getUnauthenticated(uploadUrl).body(), 0L);
    }

    /**
     * Cancels an upload session.
     *
     * @param uploadUrl the session's upload URL
     * @throws IOException if Graph returns an error
     * @throws InterruptedException if the HTTP call is interrupted
     */
    public void cancelUploadSession(String uploadUrl) throws IOException, InterruptedException {
        graph.deleteUnauthenticated(uploadUrl);
    }

    /**
     * The start of the first {@code nextExpectedRanges} entry of an upload-session
     * status.
     *
     * @param status the session status JSON
     * @param fallback the offset to return when the status names no range
     * @return the next expected offset
     */
    public static long nextExpectedOffset(JsonNode status, long fallback) {
        String range = status.path("nextExpectedRanges").path(0).asText("");
        int dash = range.indexOf('-');
        String start = dash < 0 ? range : range.substring(0, dash);
        try {
            return start.isBlank() ? fallback : Long.parseLong(start.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
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
        return GraphLinks.stripTrailingSlashes(cleaned);
    }

    private static String fileName(String destPath) {
        String normalized = normalize(destPath);
        int slash = normalized.lastIndexOf('/');
        return slash < 0 ? normalized : normalized.substring(slash + 1);
    }

    private static String encodePath(String path) {
        return GraphPaths.encodePath(normalize(path));
    }
}
