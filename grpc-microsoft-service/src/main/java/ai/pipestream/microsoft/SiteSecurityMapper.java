package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SitePrincipalType;
import ai.pipestream.microsoft.v1.SiteRoleAssignment;
import ai.pipestream.microsoft.v1.SiteRoleDefinition;
import ai.pipestream.microsoft.v1.SiteSharingLinkGroup;
import com.fasterxml.jackson.databind.JsonNode;
import java.util.UUID;

/** Maps SharePoint REST security objects to their messages. */
final class SiteSecurityMapper {

    /** Start of the login name of an Entra user. */
    static final String MEMBERSHIP_LOGIN = "i:0#.f|membership|";

    private SiteSecurityMapper() {
    }

    /** Starts the login name of the group SharePoint keeps for a sharing link. */
    static final String SHARING_LINK_LOGIN = "SharingLinks.";

    static SiteGroup toGroup(JsonNode node) {
        String login = node.path("LoginName").asText("");
        SiteGroup.Builder group = SiteGroup.newBuilder()
                .setId(node.path("Id").asInt(0))
                .setTitle(node.path("Title").asText(""))
                .setLoginName(login)
                .setDescription(node.path("Description").asText(""))
                .setOwnerTitle(node.path("OwnerTitle").asText(""))
                .setHidden(node.path("IsHiddenInUI").asBoolean(false))
                .setOnlyMembersViewMembership(node.path("OnlyAllowMembersViewMembership").asBoolean(false));
        SiteSharingLinkGroup link = sharingLink(login);
        if (link != null) {
            group.setSharingLink(link);
        }
        return group.build();
    }

    /**
     * Reads the sharing link a group login name of the form
     * {@code SharingLinks.{item}.{kind}.{share}} stands for.
     *
     * @param loginName login name of a SharePoint group
     * @return the link, or {@code null} when the name has another form
     */
    static SiteSharingLinkGroup sharingLink(String loginName) {
        if (!loginName.startsWith(SHARING_LINK_LOGIN)) {
            return null;
        }
        int kindAt = loginName.indexOf('.', SHARING_LINK_LOGIN.length());
        int shareAt = kindAt < 0 ? -1 : loginName.indexOf('.', kindAt + 1);
        if (shareAt < 0) {
            return null;
        }
        String item = loginName.substring(SHARING_LINK_LOGIN.length(), kindAt);
        String kind = loginName.substring(kindAt + 1, shareAt);
        String share = loginName.substring(shareAt + 1);
        if (item.isEmpty() || kind.isEmpty() || share.isEmpty()) {
            return null;
        }
        return SiteSharingLinkGroup.newBuilder()
                .setItemUniqueId(item)
                .setKind(kind)
                .setShareId(share)
                .setScope(sharingScope(kind))
                .build();
    }

    private static String sharingScope(String kind) {
        if (kind.startsWith("Organization")) {
            return "organization";
        }
        if (kind.startsWith("Anonymous")) {
            return "anonymous";
        }
        return kind.equals("Flexible") ? "users" : "";
    }

    /** Ends the login name of the claim for a group's owners. */
    private static final String OWNERS_SUFFIX = "_o";

    static SitePrincipal toPrincipal(JsonNode node) {
        String login = node.path("LoginName").asText("");
        SitePrincipalType type = principalType(node.path("PrincipalType").asInt(0));
        SitePrincipal.Builder principal = SitePrincipal.newBuilder()
                .setId(node.path("Id").asInt(0))
                .setPrincipalType(type)
                .setLoginName(login)
                .setTitle(node.path("Title").asText(""))
                .setEmail(node.path("Email").asText(""))
                .setUserPrincipalName(userPrincipalName(node, type, login))
                .setSiteAdmin(node.path("IsSiteAdmin").asBoolean(false))
                .setHidden(node.path("IsHiddenInUI").asBoolean(false));
        if (type == SitePrincipalType.SITE_PRINCIPAL_TYPE_SECURITY_GROUP) {
            String objectId = directoryObjectId(login);
            principal.setDirectoryObjectId(objectId)
                    .setDirectoryGroupOwners(!objectId.isEmpty() && login.endsWith(OWNERS_SUFFIX));
        }
        if (type == SitePrincipalType.SITE_PRINCIPAL_TYPE_SHAREPOINT_GROUP) {
            SiteSharingLinkGroup link = sharingLink(login);
            if (link != null) {
                principal.setSharingLink(link);
            }
        }
        return principal.build();
    }

    /** The user's principal name; when SharePoint omits it, the one in a membership login name. */
    private static String userPrincipalName(JsonNode node, SitePrincipalType type, String login) {
        String reported = node.path("UserPrincipalName").asText("");
        if (!reported.isEmpty() || type != SitePrincipalType.SITE_PRINCIPAL_TYPE_USER
                || !login.startsWith(MEMBERSHIP_LOGIN)) {
            return reported;
        }
        return login.substring(MEMBERSHIP_LOGIN.length());
    }

    static SiteRoleDefinition toRoleDefinition(JsonNode node) {
        JsonNode mask = node.path("BasePermissions");
        return SiteRoleDefinition.newBuilder()
                .setId(node.path("Id").asLong(0))
                .setName(node.path("Name").asText(""))
                .setDescription(node.path("Description").asText(""))
                .setRoleTypeKind(node.path("RoleTypeKind").asInt(0))
                .setHidden(node.path("Hidden").asBoolean(false))
                .setBasePermissionsHigh((int) mask.path("High").asLong(0))
                .setBasePermissionsLow((int) mask.path("Low").asLong(0))
                .build();
    }

    static SiteRoleAssignment toRoleAssignment(JsonNode node) {
        SiteRoleAssignment.Builder assignment = SiteRoleAssignment.newBuilder()
                .setPrincipalId(node.path("PrincipalId").asInt(0));
        if (node.path("Member").isObject()) {
            assignment.setMember(toPrincipal(node.path("Member")));
        }
        node.path("RoleDefinitionBindings")
                .forEach(role -> assignment.addRoles(toRoleDefinition(role)));
        return assignment.build();
    }

    /** SharePoint's PrincipalType numbers: 1 user, 2 distribution list, 4 security group, 8 SharePoint group. */
    static SitePrincipalType principalType(int value) {
        return switch (value) {
            case 1 -> SitePrincipalType.SITE_PRINCIPAL_TYPE_USER;
            case 2 -> SitePrincipalType.SITE_PRINCIPAL_TYPE_DISTRIBUTION_LIST;
            case 4 -> SitePrincipalType.SITE_PRINCIPAL_TYPE_SECURITY_GROUP;
            case 8 -> SitePrincipalType.SITE_PRINCIPAL_TYPE_SHAREPOINT_GROUP;
            default -> SitePrincipalType.SITE_PRINCIPAL_TYPE_UNSPECIFIED;
        };
    }

    /**
     * The Entra object id in a group's login name, such as
     * {@code c:0t.c|tenant|<id>} or {@code c:0o.c|federateddirectoryclaimprovider|<id>_o};
     * empty when the last part is not an object id.
     */
    static String directoryObjectId(String loginName) {
        String last = loginName.substring(loginName.lastIndexOf('|') + 1);
        if (last.endsWith(OWNERS_SUFFIX)) {
            last = last.substring(0, last.length() - OWNERS_SUFFIX.length());
        }
        if (last.length() != 36) {
            return "";
        }
        try {
            return UUID.fromString(last).toString();
        } catch (IllegalArgumentException e) {
            return "";
        }
    }
}
