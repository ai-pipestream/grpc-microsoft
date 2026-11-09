package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.SyncCursor;
import com.google.protobuf.InvalidProtocolBufferException;

import java.util.Base64;
import java.util.Map;

/**
 * The {@code SyncResponse.resume_cursor} codec: URL-safe, unpadded base64 of the
 * binary {@link SyncCursor}. Decoding validates every delta link inside it
 * against the configured Graph endpoint and the route it is keyed by, exactly as
 * the delta rpcs validate a client-supplied link.
 */
public final class SyncCursors {

    private SyncCursors() {
    }

    /**
     * Encodes a cursor for {@code resume_cursor}.
     *
     * @param cursor the typed cursor
     * @return URL-safe unpadded base64
     */
    public static String encode(SyncCursor cursor) {
        return Base64.getUrlEncoder().withoutPadding().encodeToString(cursor.toByteArray());
    }

    /**
     * Decodes and validates a {@code since_cursor}.
     *
     * @param encoded the client's cursor; blank yields the empty cursor
     * @param graphBaseUrl the configured Graph root every link must sit under
     * @return the typed cursor
     * @throws IllegalArgumentException when the cursor is not one this server
     *         issued, or any link in it is off-host or on the wrong route
     */
    public static SyncCursor decode(String encoded, String graphBaseUrl) {
        if (encoded == null || encoded.isBlank()) {
            return SyncCursor.getDefaultInstance();
        }
        SyncCursor cursor;
        try {
            cursor = SyncCursor.parseFrom(Base64.getUrlDecoder().decode(encoded.trim()));
        } catch (IllegalArgumentException | InvalidProtocolBufferException e) {
            throw new IllegalArgumentException("since_cursor is not a resume cursor issued by Sync");
        }
        for (Map.Entry<String, String> entry : cursor.getDriveDeltaLinksMap().entrySet()) {
            GraphLinks.requireDeltaLink(graphBaseUrl, entry.getValue(),
                    "drives", entry.getKey(), "root");
        }
        for (Map.Entry<String, String> entry : cursor.getListDeltaLinksMap().entrySet()) {
            String key = entry.getKey();
            int slash = key.lastIndexOf('/');
            if (slash <= 0 || slash == key.length() - 1) {
                throw new IllegalArgumentException("since_cursor list key " + key
                        + " is not site_id/list_id");
            }
            GraphLinks.requireDeltaLink(graphBaseUrl, entry.getValue(), "sites",
                    key.substring(0, slash), "lists", key.substring(slash + 1), "items");
        }
        if (!cursor.getSitesDeltaLink().isEmpty()) {
            GraphLinks.requireDeltaLink(graphBaseUrl, cursor.getSitesDeltaLink(), "sites");
        }
        return cursor;
    }

    /**
     * The {@link SyncCursor#getListDeltaLinksMap()} key of a list.
     *
     * @param siteId the site id
     * @param listId the list id
     * @return {@code site_id/list_id}
     */
    public static String listKey(String siteId, String listId) {
        return siteId + "/" + listId;
    }
}
