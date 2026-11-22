package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.SiteSecurable;
import com.fasterxml.jackson.databind.JsonNode;
import io.grpc.Status;
import java.io.IOException;
import java.net.URI;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Supplier;

/**
 * Reads the SharePoint REST API of a site ({@code {web}/_api}) for what Graph
 * does not expose. SharePoint accepts app-only tokens only when the app signed
 * in with a certificate.
 */
final class SharePointRest {

    /** Supplies the access tokens for one SharePoint host. */
    @FunctionalInterface
    interface Tokens {
        /**
         * The token supplier for a session and a resource.
         *
         * @param session the connection
         * @param resource scheme and host of the SharePoint site, no trailing slash
         * @return supplies a current access token
         */
        Supplier<String> forResource(MicrosoftSessions.Session session, String resource);
    }

    private static final Map<String, String> ACCEPT = Map.of("accept", "application/json;odata=nometadata");

    private final Tokens tokens;
    private final Map<String, Supplier<String>> cached = new ConcurrentHashMap<>();

    SharePointRest(Tokens tokens) {
        this.tokens = tokens;
    }

    /** Tokens from the session's certificate credential. */
    static Supplier<String> certificateTokens(MicrosoftSessions.Session session, String resource) {
        GraphAuth.Config auth = session.config().authConfig();
        if (auth.certificate() == null) {
            throw Status.FAILED_PRECONDITION.withDescription("this SharePoint call needs a certificate "
                    + "credential (" + MicrosoftConnectorConfig.ENV_CLIENT_CERT_PEM + "); SharePoint refuses "
                    + "app-only tokens issued for a client secret").asRuntimeException();
        }
        return MicrosoftSessions.tokenSupplier(new GraphAuth(auth), resource + "/.default");
    }

    /** One site's REST API. */
    final class Web {
        private final GraphClient client;

        private Web(GraphClient client) {
            this.client = client;
        }

        /** GET one object. */
        JsonNode get(String path) throws IOException, InterruptedException {
            return client.get(path, ACCEPT);
        }

        /** POST JSON and return the answer's body. */
        JsonNode post(String path, JsonNode body) throws IOException, InterruptedException {
            return client.execute(new GraphClient.Call("POST", path, ACCEPT, body)).body();
        }

        /** GET a collection, following {@code odata.nextLink}. */
        void forEach(String path, GraphPager.Visitor visitor) throws IOException, InterruptedException {
            String next = path;
            while (next != null) {
                JsonNode page = client.get(next, ACCEPT);
                for (JsonNode item : page.path("value")) {
                    if (!visitor.visit(item)) {
                        return;
                    }
                }
                String link = page.path("odata.nextLink").asText("");
                next = link.isEmpty() ? null : link.startsWith("http") ? link : "/" + link;
            }
        }
    }

    /**
     * The REST API of the site with this Graph site id.
     *
     * @param session the connection
     * @param siteId Graph site id
     * @return the site's REST API
     */
    Web web(MicrosoftSessions.Session session, String siteId) throws IOException, InterruptedException {
        if (siteId.isBlank()) {
            throw new IllegalArgumentException("site_id is required");
        }
        String webUrl = session.graph().get("/sites/" + GraphPaths.site(siteId) + "?$select=id,webUrl")
                .path("webUrl").asText("");
        if (webUrl.isEmpty()) {
            throw Status.NOT_FOUND.withDescription("site " + siteId + " has no web URL").asRuntimeException();
        }
        URI uri = URI.create(webUrl);
        String resource = uri.getScheme() + "://" + uri.getAuthority();
        Supplier<String> bearer = cached.computeIfAbsent(session.connectionId() + "\n" + resource,
                key -> tokens.forResource(session, resource));
        return new Web(session.graph().sibling(GraphLinks.stripTrailingSlashes(webUrl) + "/_api", bearer));
    }

    /** The site id a securable belongs to. */
    static String siteId(SiteSecurable securable) {
        return switch (securable.getSecurableCase()) {
            case SITE_ID -> securable.getSiteId();
            case LIST -> securable.getList().getSiteId();
            case LIST_ITEM -> securable.getListItem().getSiteId();
            case SECURABLE_NOT_SET -> throw new IllegalArgumentException("securable is required");
        };
    }

    /** The REST path of a securable, relative to {@code {web}/_api}. */
    static String path(SiteSecurable securable) {
        return switch (securable.getSecurableCase()) {
            case SITE_ID -> "/web";
            case LIST -> list(securable.getList().getListId());
            case LIST_ITEM -> {
                ListItemRef item = securable.getListItem();
                yield list(item.getListId()) + "/items(" + itemId(item.getItemId()) + ")";
            }
            case SECURABLE_NOT_SET -> throw new IllegalArgumentException("securable is required");
        };
    }

    /**
     * The REST path of the columns of a list or of the site.
     *
     * @param listId the list's GUID; empty = the site's columns
     * @return the path, relative to {@code {web}/_api}
     */
    static String fields(String listId) {
        return (listId.isBlank() ? "/web" : list(listId)) + "/fields";
    }

    private static String list(String listId) {
        try {
            return "/web/lists(guid'" + UUID.fromString(listId.trim()) + "')";
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("list_id must be the list's GUID", e);
        }
    }

    private static int itemId(String itemId) {
        try {
            return Integer.parseInt(itemId.trim());
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("item_id must be the list item's number", e);
        }
    }
}
