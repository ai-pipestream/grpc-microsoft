package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.SiteAcl;
import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SiteGroupMembers;
import ai.pipestream.microsoft.v1.SiteSecurable;
import ai.pipestream.sync.v1.AclCompleteness;
import com.fasterxml.jackson.databind.JsonNode;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Reads the SharePoint security of one site for a sync pass: its groups with
 * their members, the role assignments of the site and of every list and list
 * item that has its own, and which of those applies to each list item.
 */
final class SiteSecurityReader {

    private static final String ASSIGNMENTS = "/roleassignments?$expand=Member,RoleDefinitionBindings";
    private static final String ITEMS =
            "/items?$select=Id,HasUniqueRoleAssignments,FileRef,FileDirRef&$top=1000";

    /**
     * The security of one list.
     *
     * @param listOwner entity id of the role assignments that apply to the list
     * @param itemOwners list item id to the entity id of the role assignments
     *        that apply to it, for the items where that is not {@code listOwner}
     * @param acls the role assignments of the list and its items that have their own
     */
    record ListSecurity(String listOwner, Map<String, String> itemOwners, List<SiteAcl> acls) {

        /** Entity id of the role assignments that apply to a list item. */
        String owner(String listItemId) {
            return itemOwners.getOrDefault(listItemId, listOwner);
        }
    }

    private final SharePointRest.Web web;
    private final String siteId;

    SiteSecurityReader(SharePointRest.Web web, String siteId) {
        this.web = web;
        this.siteId = siteId;
    }

    /** Every SharePoint group of the site with its members. */
    List<SiteGroupMembers> groups() throws IOException, InterruptedException {
        List<SiteGroup> groups = new ArrayList<>();
        web.forEach("/web/sitegroups", node -> groups.add(SiteSecurityMapper.toGroup(node)));
        List<SiteGroupMembers> out = new ArrayList<>();
        for (SiteGroup group : groups) {
            SiteGroupMembers.Builder members = SiteGroupMembers.newBuilder().setSiteId(siteId).setGroup(group);
            try {
                web.forEach("/web/sitegroups/getbyid(" + group.getId() + ")/users", node -> {
                    members.addMembers(SiteSecurityMapper.toPrincipal(node));
                    return true;
                });
                members.setCompleteness(AclCompleteness.ACL_COMPLETENESS_COMPLETE);
            } catch (GraphClient.GraphApiException e) {
                members.clearMembers().setCompleteness(AclCompleteness.ACL_COMPLETENESS_UNAVAILABLE);
            }
            out.add(members.build());
        }
        return out;
    }

    /** The role assignments of the site itself. */
    SiteAcl site() throws IOException, InterruptedException {
        return acl(SiteSecurable.newBuilder().setSiteId(siteId).build());
    }

    /** The security of one list and its items. */
    ListSecurity list(String listId) throws IOException, InterruptedException {
        SiteSecurable list = SiteSecurable.newBuilder()
                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId)).build();
        String listPath = SharePointRest.path(list);
        List<SiteAcl> acls = new ArrayList<>();
        String listOwner;
        Map<String, String> itemOwners = new HashMap<>();
        try {
            boolean unique = web.get(listPath + "?$select=HasUniqueRoleAssignments")
                    .path("HasUniqueRoleAssignments").asBoolean(false);
            if (unique) {
                acls.add(acl(list));
                listOwner = MicrosoftEntityIds.siteAcl(list);
            } else {
                listOwner = MicrosoftEntityIds.siteAcl(SiteSecurable.newBuilder().setSiteId(siteId).build());
            }
            Map<String, String> folderOf = new HashMap<>();
            Map<String, String> ownerByPath = new HashMap<>();
            web.forEach(listPath + ITEMS, node -> {
                String id = node.path("Id").asText("");
                folderOf.put(id, node.path("FileDirRef").asText(""));
                if (node.path("HasUniqueRoleAssignments").asBoolean(false)) {
                    SiteSecurable item = SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder()
                            .setSiteId(siteId).setListId(listId).setItemId(id)).build();
                    acls.add(acl(item));
                    String owner = MicrosoftEntityIds.siteAcl(item);
                    itemOwners.put(id, owner);
                    ownerByPath.put(node.path("FileRef").asText(""), owner);
                }
                return true;
            });
            if (!ownerByPath.isEmpty()) {
                folderOf.forEach((id, folder) -> {
                    if (!itemOwners.containsKey(id)) {
                        String owner = nearest(ownerByPath, folder);
                        if (owner != null) {
                            itemOwners.put(id, owner);
                        }
                    }
                });
            }
        } catch (GraphClient.GraphApiException e) {
            // Nothing of the list can be trusted: everything in it points at one unavailable record.
            acls.clear();
            itemOwners.clear();
            acls.add(SiteAcl.newBuilder().setSecurable(list)
                    .setCompleteness(AclCompleteness.ACL_COMPLETENESS_UNAVAILABLE).build());
            listOwner = MicrosoftEntityIds.siteAcl(list);
        }
        return new ListSecurity(listOwner, Map.copyOf(itemOwners), List.copyOf(acls));
    }

    /** The owner of the nearest folder at or above {@code folder} that has one. */
    private static String nearest(Map<String, String> ownerByPath, String folder) {
        String path = folder;
        while (!path.isEmpty()) {
            String owner = ownerByPath.get(path);
            if (owner != null) {
                return owner;
            }
            int slash = path.lastIndexOf('/');
            path = slash <= 0 ? "" : path.substring(0, slash);
        }
        return null;
    }

    private SiteAcl acl(SiteSecurable securable) throws IOException, InterruptedException {
        SiteAcl.Builder acl = SiteAcl.newBuilder().setSecurable(securable);
        try {
            web.forEach(SharePointRest.path(securable) + ASSIGNMENTS, node -> {
                acl.addAssignments(SiteSecurityMapper.toRoleAssignment(node));
                return true;
            });
            acl.setCompleteness(AclCompleteness.ACL_COMPLETENESS_COMPLETE);
        } catch (GraphClient.GraphApiException e) {
            acl.clearAssignments().setCompleteness(AclCompleteness.ACL_COMPLETENESS_UNAVAILABLE);
        }
        return acl.build();
    }
}
