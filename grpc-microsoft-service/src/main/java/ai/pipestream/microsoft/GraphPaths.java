package ai.pipestream.microsoft;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

/**
 * Graph v1.0 route and query builders for the {@code MicrosoftService} rpcs, plus
 * the {@code $select} lists that make Graph return every property the contract
 * models. Graph omits {@code sharepointIds}, {@code publication}, the drive and
 * list {@code system} facet and several user properties unless they are
 * selected, so every entity read names its properties explicitly. Ids are
 * percent-encoded per path segment so a request field can never add a query
 * string or another path segment to the route.
 */
public final class GraphPaths {

    /**
     * Every {@code driveItem} property the contract maps, plus the pre-authenticated
     * download URL annotation (Graph documents {@code select=id,@microsoft.graph.downloadUrl}).
     */
    public static final String DRIVE_ITEM_SELECT = "id,name,eTag,cTag,createdBy,createdDateTime,"
            + "lastModifiedBy,lastModifiedDateTime,parentReference,webUrl,webDavUrl,description,"
            + "size,file,folder,fileSystemInfo,deleted,root,package,remoteItem,shared,"
            + "specialFolder,image,photo,video,audio,location,malware,sharepointIds,"
            + "pendingOperations,@microsoft.graph.downloadUrl";
    /** The extra {@code driveItem} property only returned when selected. */
    public static final String PUBLICATION = "publication";
    /** Every {@code drive} property the contract maps. */
    public static final String DRIVE_SELECT = "id,name,driveType,description,webUrl,eTag,"
            + "createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,owner,quota,"
            + "sharePointIds,system";
    /** Every {@code site} property the contract maps. */
    public static final String SITE_SELECT = "id,name,displayName,description,webUrl,eTag,"
            + "createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,isPersonalSite,root,"
            + "siteCollection,sharepointIds,parentReference";
    /** Every {@code list} property the contract maps. */
    public static final String LIST_SELECT = "id,name,displayName,description,webUrl,eTag,"
            + "createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,list,sharepointIds,"
            + "system,parentReference";
    /** Every {@code listItem} property the contract maps (fields come from $expand). */
    public static final String LIST_ITEM_SELECT = "id,name,webUrl,description,eTag,"
            + "createdDateTime,lastModifiedDateTime,createdBy,lastModifiedBy,contentType,"
            + "sharepointIds,parentReference,deleted";
    /** Every {@code user} property {@code GraphUser} maps. */
    public static final String USER_SELECT = "id,displayName,userPrincipalName,mail,givenName,"
            + "surname,jobTitle,department,officeLocation,userType,accountEnabled,"
            + "preferredLanguage";
    /** Graph preference required for $filter on non-indexed list columns. */
    public static final String HONOR_NON_INDEXED = "HonorNonIndexedQueriesWarningMayFailRandomly";

    private GraphPaths() {
    }

    /**
     * Percent-encodes one id for use as a single path segment. Characters Graph
     * uses inside ids ({@code ! , : - _ . ~}) stay literal; {@code /}, {@code ?},
     * {@code #}, spaces and everything else are encoded.
     *
     * @param id the id; must not be blank
     * @param field the request field name, for the error message
     * @return the encoded segment
     * @throws IllegalArgumentException when {@code id} is blank or a dot segment
     */
    public static String segment(String id, String field) {
        if (id == null || id.isBlank()) {
            throw new IllegalArgumentException(field + " is required");
        }
        if (id.equals(".") || id.equals("..")) {
            throw new IllegalArgumentException(field + " is not a valid id");
        }
        StringBuilder out = new StringBuilder(id.length());
        for (int i = 0; i < id.length(); i++) {
            char c = id.charAt(i);
            if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')
                    || c == '-' || c == '_' || c == '.' || c == '~' || c == '!' || c == ','
                    || c == ':' || c == '@' || c == '$' || c == '\'' || c == '(' || c == ')') {
                out.append(c);
            } else {
                out.append(URLEncoder.encode(String.valueOf(c), StandardCharsets.UTF_8)
                        .replace("+", "%20"));
            }
        }
        return out.toString();
    }

    /**
     * The site id segment: either a composite id ({@code host,collection,web}) or
     * Graph's {@code {hostname}:/{server-relative-path}} form, whose path keeps
     * its slashes (each path segment is encoded on its own).
     *
     * @param siteId the request's site id
     * @return the encoded site segment
     */
    public static String site(String siteId) {
        if (siteId != null) {
            int colon = siteId.indexOf(":/");
            if (colon > 0) {
                String host = segment(siteId.substring(0, colon), "site_id");
                String path = siteId.substring(colon + 1);
                return host + ":" + encodePath(path);
            }
        }
        return segment(siteId, "site_id");
    }

    /**
     * Percent-encodes a slash-separated path one segment at a time; backslashes
     * become slashes, empty and dot segments are dropped.
     *
     * @param path a drive-relative path such as {@code /Shared/report.docx}
     * @return {@code /seg/seg}, or empty for the root
     */
    public static String encodePath(String path) {
        StringBuilder encoded = new StringBuilder();
        if (path == null) {
            return "";
        }
        String normalized = path.replace('\\', '/');
        int start = 0;
        while (start <= normalized.length()) {
            int slash = normalized.indexOf('/', start);
            int end = slash < 0 ? normalized.length() : slash;
            String part = normalized.substring(start, end);
            start = end + 1;
            if (part.isEmpty() || part.equals(".")) {
                continue;
            }
            if (part.equals("..")) {
                throw new IllegalArgumentException("path must not contain '..' segments");
            }
            encoded.append('/').append(URLEncoder.encode(part, StandardCharsets.UTF_8)
                    .replace("+", "%20"));
        }
        return encoded.toString();
    }

    /**
     * Encodes one query parameter value.
     *
     * @param value the raw value
     * @return the encoded value
     */
    public static String query(String value) {
        return URLEncoder.encode(value, StandardCharsets.UTF_8).replace("+", "%20");
    }

    /**
     * {@code /drives/{drive-id}}.
     *
     * @param driveId the drive id
     * @return the route
     */
    public static String drive(String driveId) {
        return "/drives/" + segment(driveId, "drive_id");
    }

    /**
     * {@code /drives/{drive-id}/items/{item-id}}.
     *
     * @param driveId the drive id
     * @param itemId the item id
     * @return the route
     */
    public static String item(String driveId, String itemId) {
        return drive(driveId) + "/items/" + segment(itemId, "item_id");
    }

    /**
     * {@code /drives/{drive-id}/root:/{path}:}, or {@code /drives/{drive-id}/root}
     * for an empty path.
     *
     * @param driveId the drive id
     * @param path the drive-relative path
     * @return the route
     */
    public static String itemByPath(String driveId, String path) {
        String encoded = encodePath(path);
        return encoded.isEmpty() ? drive(driveId) + "/root" : drive(driveId) + "/root:" + encoded + ":";
    }

    /**
     * {@code /sites/{site-id}/lists/{list-id}}.
     *
     * @param siteId the site id
     * @param listId the list id
     * @return the route
     */
    public static String list(String siteId, String listId) {
        return "/sites/" + site(siteId) + "/lists/" + segment(listId, "list_id");
    }

    /**
     * {@code /sites/{site-id}/lists/{list-id}/items/{item-id}}.
     *
     * @param siteId the site id
     * @param listId the list id
     * @param itemId the list item id
     * @return the route
     */
    public static String listItem(String siteId, String listId, String itemId) {
        return list(siteId, listId) + "/items/" + segment(itemId, "item_id");
    }

    /**
     * The drive-item {@code $select} for a request.
     *
     * @param publication whether to add the {@code publication} facet
     * @return the select list
     */
    public static String driveItemSelect(boolean publication) {
        return publication ? DRIVE_ITEM_SELECT + "," + PUBLICATION : DRIVE_ITEM_SELECT;
    }
}
