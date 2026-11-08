package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.PermissionTarget;
import com.fasterxml.jackson.databind.JsonNode;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Reads list items of one list with their typed column values and the requested
 * expansions. A record that arrives without {@code fields} (a delta page whose
 * link dropped the {@code $expand}, or an endpoint that ignores it) is re-read
 * with {@code GET .../items/{id}?$expand=fields}, so column values are never
 * silently missing; deleted records are passed through identity-only.
 */
final class ListItemReader {

    private final GraphPager pager;
    private final MicrosoftMapper mapper;
    private final String siteId;
    private final String listId;
    private final boolean driveItem;
    private final boolean permissions;

    ListItemReader(GraphPager pager, MicrosoftMapper mapper, String siteId, String listId,
            boolean driveItem, boolean permissions) {
        this.pager = pager;
        this.mapper = mapper;
        this.siteId = siteId;
        this.listId = listId;
        this.driveItem = driveItem;
        this.permissions = permissions;
    }

    String expand() {
        return driveItem ? "fields,driveItem" : "fields";
    }

    /** Candidate URLs for {@code /items} (preferred first). */
    List<String> listCandidates(String filter) {
        String base = GraphPaths.list(siteId, listId) + "/items";
        String filterQuery = filter == null || filter.isBlank() ? ""
                : "&$filter=" + GraphPaths.query(filter);
        return List.of(base + "?$select=" + GraphPaths.LIST_ITEM_SELECT + "&$expand=" + expand()
                        + filterQuery,
                base + "?$expand=" + expand() + filterQuery);
    }

    /** Candidate URLs for {@code /items/{id}}. */
    List<String> itemCandidates(String itemId) {
        String base = GraphPaths.listItem(siteId, listId, itemId);
        return List.of(base + "?$select=" + GraphPaths.LIST_ITEM_SELECT + "&$expand=" + expand(),
                base + "?$expand=" + expand());
    }

    /** Candidate URLs for a fresh {@code /items/delta}. */
    List<String> deltaCandidates() {
        String base = GraphPaths.list(siteId, listId) + "/items/delta";
        return List.of(base + "?$select=" + GraphPaths.LIST_ITEM_SELECT + "&$expand=" + expand(),
                base + "?$expand=" + expand(),
                base);
    }

    /** Maps one record, re-reading it when Graph left out the expansions. */
    ListItem read(JsonNode node) throws IOException, InterruptedException {
        JsonNode source = node;
        boolean deleted = node.path("deleted").isObject() || node.path("@removed").isObject();
        if (!deleted && (!node.path("fields").isObject()
                || (driveItem && !node.path("driveItem").isObject()))
                && !node.path("id").asText("").isEmpty()) {
            source = pager.get(itemCandidates(node.path("id").asText()), Map.of());
        }
        ListItem item = mapper.toListItem(source, siteId, listId);
        if (permissions && !deleted) {
            PermissionTarget target = PermissionTarget.newBuilder()
                    .setListItem(ListItemRef.newBuilder()
                            .setSiteId(siteId).setListId(listId).setItemId(item.getId()))
                    .build();
            ListItem.Builder b = item.toBuilder();
            pager.forEach(List.of(GraphPaths.listItem(siteId, listId, item.getId()) + "/permissions"),
                    Map.of(), p -> {
                        b.addPermissions(mapper.toPermission(p, target));
                        return true;
                    });
            item = b.build();
        }
        return item;
    }
}
