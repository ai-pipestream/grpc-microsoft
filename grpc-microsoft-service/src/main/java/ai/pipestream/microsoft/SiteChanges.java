package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListSiteChangesRequest;
import ai.pipestream.microsoft.v1.SiteChange;
import ai.pipestream.microsoft.v1.SiteChangeObject;
import ai.pipestream.microsoft.v1.SiteChangeScope;
import ai.pipestream.microsoft.v1.SiteChangeType;
import ai.pipestream.microsoft.v1.SiteSecurable;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;
import java.io.IOException;
import java.util.EnumMap;
import java.util.EnumSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;

/** Reads a SharePoint change log ({@code getchanges}) and maps its entries. */
final class SiteChanges {

    /** The change token that asks for the end of the log and no entries. */
    static final String LATEST = "latest";

    /** In the error code SharePoint answers, with HTTP 500, to a token it cannot parse. */
    private static final String MALFORMED_TOKEN = "System.FormatException";

    /** The ChangeQuery flag that asks for each kind of object. */
    private static final Map<SiteChangeObject, String> OBJECT_FLAGS = objectFlags();
    /** The ChangeQuery flag that asks for each type of change; some types have none. */
    private static final Map<SiteChangeType, String> TYPE_FLAGS = typeFlags();
    /** Every ChangeQuery flag that names a type of change. */
    private static final List<String> ALL_TYPE_FLAGS = List.of("Add", "Update", "DeleteObject", "Rename",
            "Move", "Restore", "RoleDefinitionAdd", "RoleDefinitionDelete", "RoleDefinitionUpdate",
            "RoleAssignmentAdd", "RoleAssignmentDelete", "GroupMembershipAdd", "GroupMembershipDelete",
            "SystemUpdate", "Navigation");

    private SiteChanges() {
    }

    /**
     * Reads the change log from the request's token to its end.
     *
     * @param web the site's REST API
     * @param request names the scope, the token and the filter
     * @param out receives each entry that passes the filter, oldest first
     * @return the change token to resume from
     * @throws IllegalArgumentException if SharePoint cannot parse the token
     */
    static String read(SharePointRest.Web web, ListSiteChangesRequest request, Consumer<SiteChange> out)
            throws IOException, InterruptedException {
        String scope = scopePath(request);
        String token = request.getChangeToken().trim();
        if (token.equals(LATEST)) {
            return currentToken(web, scope);
        }
        Set<SiteChangeObject> objects = set(SiteChangeObject.class, request.getObjectsList(),
                SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED, SiteChangeObject.UNRECOGNIZED);
        Set<SiteChangeType> types = set(SiteChangeType.class, request.getTypesList(),
                SiteChangeType.SITE_CHANGE_TYPE_UNSPECIFIED, SiteChangeType.UNRECOGNIZED);
        while (true) {
            JsonNode page;
            try {
                page = web.postTyped(scope + "/getchanges", query(objects, types, token));
            } catch (GraphClient.GraphApiException e) {
                if (!token.isEmpty() && e.code().contains(MALFORMED_TOKEN)) {
                    throw new IllegalArgumentException("change_token is not a change token of SharePoint");
                }
                throw e;
            }
            String last = "";
            for (JsonNode node : page.path("value")) {
                SiteChange change = toChange(node, web.host(), request.getSiteId());
                last = change.getChangeToken();
                if ((objects.isEmpty() || objects.contains(change.getObject()))
                        && (types.isEmpty() || types.contains(change.getType()))) {
                    out.accept(change);
                }
            }
            if (last.isEmpty() || last.equals(token)) {
                return token.isEmpty() ? currentToken(web, scope) : token;
            }
            token = last;
        }
    }

    /** The REST path of the request's change log, relative to {@code {web}/_api}. */
    static String scopePath(ListSiteChangesRequest request) {
        boolean list = request.getScope() == SiteChangeScope.SITE_CHANGE_SCOPE_LIST;
        if (list == request.getListId().isBlank()) {
            throw new IllegalArgumentException(list ? "list_id is required for the list scope"
                    : "list_id is only allowed with the list scope");
        }
        return switch (request.getScope()) {
            case SITE_CHANGE_SCOPE_LIST -> SharePointRest.listPath(request.getListId());
            case SITE_CHANGE_SCOPE_SITE_COLLECTION -> "/site";
            case SITE_CHANGE_SCOPE_UNSPECIFIED, SITE_CHANGE_SCOPE_WEB -> "/web";
            case UNRECOGNIZED -> throw new IllegalArgumentException("unknown scope");
        };
    }

    /**
     * The {@code getchanges} request body.
     *
     * @param objects the kinds of object asked for; empty = all
     * @param types the types of change asked for; empty = all
     * @param token the change token to start after; empty = the start of the log
     * @return the request JSON
     */
    static ObjectNode query(Set<SiteChangeObject> objects, Set<SiteChangeType> types, String token) {
        ObjectNode body = GraphClient.object();
        ObjectNode query = body.putObject("query");
        OBJECT_FLAGS.forEach((object, flag) -> {
            if (objects.isEmpty() || objects.contains(object)) {
                query.put(flag, true);
            }
        });
        // A type SharePoint has no flag for comes only when every flag is set.
        boolean all = types.isEmpty() || !TYPE_FLAGS.keySet().containsAll(types);
        for (String flag : ALL_TYPE_FLAGS) {
            if (all || types.stream().anyMatch(type -> flag.equals(TYPE_FLAGS.get(type)))) {
                query.put(flag, true);
            }
        }
        if (!token.isEmpty()) {
            query.putObject("ChangeTokenStart").put("StringValue", token);
        }
        return body;
    }

    /**
     * Maps one entry of the change log.
     *
     * @param node the entry
     * @param host host name of the site
     * @param siteId the site id of the request
     * @return the change
     */
    static SiteChange toChange(JsonNode node, String host, String siteId) {
        String odataType = node.path("odata.type").asText("");
        SiteChangeObject object = object(odataType);
        SiteChangeType type = SiteChangeType.forNumber(node.path("ChangeType").asInt(0));
        String collection = guid(node, "SiteId");
        String web = guid(node, "WebId");
        String list = guid(node, "ListId");
        String item = node.path("ItemId").isNumber() ? node.path("ItemId").asText() : "";
        SiteChange.Builder change = SiteChange.newBuilder()
                .setObject(object)
                .setType(type == null ? SiteChangeType.SITE_CHANGE_TYPE_UNSPECIFIED : type)
                .setChangeToken(node.path("ChangeToken").path("StringValue").asText(""))
                .setSiteCollectionId(collection)
                .setWebId(web)
                .setListId(list)
                .setItemId(item)
                .setUniqueId(guid(node, "UniqueId"))
                .setGroupId(node.path("GroupId").asInt(0))
                .setUserId(node.path("UserId").asInt(0))
                .setFieldId(guid(node, "FieldId"))
                .setContentTypeId(node.path("ContentTypeId").path("StringValue").asText(""))
                .setViewId(guid(node, "ViewId"))
                .setAlertId(guid(node, "AlertId"));
        Timestamp time = MicrosoftMapper.timestamp(node.path("Time").asText(""));
        if (time != null) {
            change.setTime(time);
        }
        if (object == SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED && !odataType.isEmpty()) {
            change.putMetadata("odata.type", odataType);
        }
        if (type == null) {
            change.putMetadata("ChangeType", node.path("ChangeType").asText(""));
        }
        if (!collection.isEmpty() && !web.isEmpty()) {
            String site = host + "," + collection + "," + web;
            change.setSiteId(site);
            // The request's own spelling of the site id when it names the same site.
            String named = siteId.equalsIgnoreCase(site) ? siteId : site;
            switch (object) {
                case SITE_CHANGE_OBJECT_WEB ->
                        change.setSecurable(SiteSecurable.newBuilder().setSiteId(named));
                case SITE_CHANGE_OBJECT_LIST -> change.setSecurable(SiteSecurable.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(named).setListId(list)));
                case SITE_CHANGE_OBJECT_ITEM -> change.setSecurable(SiteSecurable.newBuilder()
                        .setListItem(ListItemRef.newBuilder().setSiteId(named).setListId(list)
                                .setItemId(item)));
                default -> {
                }
            }
        }
        return change.build();
    }

    private static String currentToken(SharePointRest.Web web, String scope)
            throws IOException, InterruptedException {
        return web.get(scope + "?$select=CurrentChangeToken").path("CurrentChangeToken")
                .path("StringValue").asText("");
    }

    private static SiteChangeObject object(String odataType) {
        return switch (odataType) {
            case "SP.ChangeItem" -> SiteChangeObject.SITE_CHANGE_OBJECT_ITEM;
            case "SP.ChangeList" -> SiteChangeObject.SITE_CHANGE_OBJECT_LIST;
            case "SP.ChangeWeb" -> SiteChangeObject.SITE_CHANGE_OBJECT_WEB;
            case "SP.ChangeSite" -> SiteChangeObject.SITE_CHANGE_OBJECT_SITE_COLLECTION;
            case "SP.ChangeFile" -> SiteChangeObject.SITE_CHANGE_OBJECT_FILE;
            case "SP.ChangeFolder" -> SiteChangeObject.SITE_CHANGE_OBJECT_FOLDER;
            case "SP.ChangeUser" -> SiteChangeObject.SITE_CHANGE_OBJECT_USER;
            case "SP.ChangeGroup" -> SiteChangeObject.SITE_CHANGE_OBJECT_GROUP;
            case "SP.ChangeField" -> SiteChangeObject.SITE_CHANGE_OBJECT_FIELD;
            case "SP.ChangeContentType" -> SiteChangeObject.SITE_CHANGE_OBJECT_CONTENT_TYPE;
            case "SP.ChangeView" -> SiteChangeObject.SITE_CHANGE_OBJECT_VIEW;
            case "SP.ChangeAlert" -> SiteChangeObject.SITE_CHANGE_OBJECT_ALERT;
            case "SP.ChangeSecurityPolicy" -> SiteChangeObject.SITE_CHANGE_OBJECT_SECURITY_POLICY;
            default -> SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED;
        };
    }

    /** A GUID property in lower case; empty for a missing or all-zero one. */
    private static String guid(JsonNode node, String name) {
        String value = node.path(name).asText("").toLowerCase(java.util.Locale.ROOT);
        return value.equals("00000000-0000-0000-0000-000000000000") ? "" : value;
    }

    private static <E extends Enum<E>> Set<E> set(Class<E> type, List<E> values, E unspecified,
            E unrecognized) {
        Set<E> set = EnumSet.noneOf(type);
        for (E value : values) {
            if (value == unspecified || value == unrecognized) {
                throw new IllegalArgumentException("a filter value is not set");
            }
            set.add(value);
        }
        return set;
    }

    private static Map<SiteChangeObject, String> objectFlags() {
        Map<SiteChangeObject, String> flags = new EnumMap<>(SiteChangeObject.class);
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_ITEM, "Item");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_LIST, "List");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_WEB, "Web");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_SITE_COLLECTION, "Site");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_FILE, "File");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_FOLDER, "Folder");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_USER, "User");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_GROUP, "Group");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_FIELD, "Field");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_CONTENT_TYPE, "ContentType");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_VIEW, "View");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_ALERT, "Alert");
        flags.put(SiteChangeObject.SITE_CHANGE_OBJECT_SECURITY_POLICY, "SecurityPolicy");
        return flags;
    }

    private static Map<SiteChangeType, String> typeFlags() {
        Map<SiteChangeType, String> flags = new EnumMap<>(SiteChangeType.class);
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ADD, "Add");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_UPDATE, "Update");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_DELETE, "DeleteObject");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_RENAME, "Rename");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_MOVE_AWAY, "Move");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_MOVE_INTO, "Move");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_RESTORE, "Restore");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ROLE_DEFINITION_ADD, "RoleDefinitionAdd");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ROLE_DEFINITION_DELETE, "RoleDefinitionDelete");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ROLE_DEFINITION_UPDATE, "RoleDefinitionUpdate");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ROLE_ASSIGNMENT_ADD, "RoleAssignmentAdd");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_ROLE_ASSIGNMENT_DELETE, "RoleAssignmentDelete");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_ADD, "GroupMembershipAdd");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_DELETE, "GroupMembershipDelete");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_SYSTEM_UPDATE, "SystemUpdate");
        flags.put(SiteChangeType.SITE_CHANGE_TYPE_NAVIGATION, "Navigation");
        return flags;
    }
}
