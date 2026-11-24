package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AddSiteGroupMemberRequest;
import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SiteSecurable;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import java.io.IOException;

/** Changes the SharePoint groups, group members and role assignments of a site through its REST API. */
final class SiteSecurityWriter {

    private static final JsonNodeFactory JSON = JsonNodeFactory.instance;

    private SiteSecurityWriter() {
    }

    /**
     * Creates a SharePoint group.
     *
     * @param web the site's REST API
     * @param title name of the group
     * @param description description; empty = none
     * @return the new group
     */
    static SiteGroup createGroup(SharePointRest.Web web, String title, String description)
            throws IOException, InterruptedException {
        if (title.isBlank()) {
            throw new IllegalArgumentException("title is required");
        }
        ObjectNode body = JSON.objectNode().put("Title", title.trim());
        if (!description.isEmpty()) {
            body.put("Description", description);
        }
        return SiteSecurityMapper.toGroup(web.post("/web/sitegroups", body));
    }

    /**
     * Deletes a SharePoint group.
     *
     * @param web the site's REST API
     * @param groupId id of the group
     */
    static void deleteGroup(SharePointRest.Web web, int groupId) throws IOException, InterruptedException {
        web.post("/web/sitegroups/removebyid(" + id(groupId, "group_id") + ")", JSON.objectNode());
    }

    /**
     * Adds a member to a SharePoint group.
     *
     * @param web the site's REST API
     * @param request names the group and the member
     * @return the member as the site knows it
     */
    static SitePrincipal addMember(SharePointRest.Web web, AddSiteGroupMemberRequest request)
            throws IOException, InterruptedException {
        String login = switch (request.getMemberCase()) {
            case LOGIN_NAME -> request.getLoginName().trim();
            case USER_PRINCIPAL_NAME -> request.getUserPrincipalName().isBlank() ? ""
                    : SiteSecurityMapper.MEMBERSHIP_LOGIN + request.getUserPrincipalName().trim();
            case MEMBER_NOT_SET -> "";
        };
        if (login.isEmpty()) {
            throw new IllegalArgumentException("login_name or user_principal_name is required");
        }
        return SiteSecurityMapper.toPrincipal(web.post(members(request.getGroupId()),
                JSON.objectNode().put("LoginName", login)));
    }

    /**
     * Removes a member from a SharePoint group.
     *
     * @param web the site's REST API
     * @param groupId id of the group
     * @param memberId id of the member within the site collection
     */
    static void removeMember(SharePointRest.Web web, int groupId, int memberId)
            throws IOException, InterruptedException {
        web.post(members(groupId) + "/removebyid(" + id(memberId, "member_id") + ")", JSON.objectNode());
    }

    /**
     * Grants or takes one permission level on a securable.
     *
     * @param web the site's REST API
     * @param securable the site, list or list item
     * @param principalId id of the principal within the site collection
     * @param roleDefinitionId id of the permission level
     * @param grant true = grant, false = take away
     */
    static void roleAssignment(SharePointRest.Web web, SiteSecurable securable, int principalId,
            long roleDefinitionId, boolean grant) throws IOException, InterruptedException {
        if (roleDefinitionId <= 0) {
            throw new IllegalArgumentException("role_definition_id is required");
        }
        web.post(SharePointRest.path(securable) + "/roleassignments/"
                + (grant ? "addroleassignment" : "removeroleassignment")
                + "(principalid=" + id(principalId, "principal_id") + ",roledefid=" + roleDefinitionId + ")",
                JSON.objectNode());
    }

    private static String members(int groupId) {
        return "/web/sitegroups/getbyid(" + id(groupId, "group_id") + ")/users";
    }

    private static int id(int value, String field) {
        if (value <= 0) {
            throw new IllegalArgumentException(field + " is required");
        }
        return value;
    }
}
