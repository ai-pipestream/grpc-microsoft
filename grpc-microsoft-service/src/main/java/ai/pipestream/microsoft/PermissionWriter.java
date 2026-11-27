package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.CreateSharingLinkRequest;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.GrantLinkAccessRequest;
import ai.pipestream.microsoft.v1.InviteRecipientsRequest;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.RevokeLinkAccessRequest;
import ai.pipestream.microsoft.v1.SharingRecipient;
import ai.pipestream.microsoft.v1.UpdatePermissionRequest;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;
import java.io.IOException;
import java.time.Instant;
import java.util.List;
import java.util.Map;

/** Creates sharing links and grants, changes and removes permissions through Graph. */
final class PermissionWriter {

    private PermissionWriter() {
    }

    /**
     * {@code {resource}/permissions} of a drive item, list item, list or site.
     *
     * @param target the resource
     * @return the Graph path
     */
    static String permissions(PermissionTarget target) {
        return resource(target, true) + "/permissions";
    }

    /**
     * {@code {resource}/permissions/{permission-id}}.
     *
     * @param target the resource
     * @param permissionId id of the permission
     * @return the Graph path
     */
    static String permission(PermissionTarget target, String permissionId) {
        return permissions(target) + "/" + GraphPaths.segment(permissionId, "permission_id");
    }

    /**
     * Creates a sharing link.
     *
     * @param graph the Graph API
     * @param request the item and the link to make
     * @return Graph's answer; status 201 when the link is new
     */
    static GraphClient.Reply createLink(GraphClient graph, CreateSharingLinkRequest request)
            throws IOException, InterruptedException {
        if (request.getType().isBlank()) {
            throw new IllegalArgumentException("type is required");
        }
        if (request.getSendNotification() && request.getRecipientsCount() == 0) {
            throw new IllegalArgumentException("send_notification needs recipients");
        }
        ObjectNode body = GraphClient.object().put("type", request.getType().trim());
        if (!request.getScope().isBlank()) {
            body.put("scope", request.getScope().trim());
        }
        if (request.hasExpiresAt()) {
            body.put("expirationDateTime", instant(request.getExpiresAt()));
        }
        if (!request.getPassword().isEmpty()) {
            body.put("password", request.getPassword());
        }
        if (request.hasRetainInheritedPermissions()) {
            body.put("retainInheritedPermissions", request.getRetainInheritedPermissions());
        }
        if (request.getRecipientsCount() > 0) {
            body.set("recipients", recipients(request.getRecipientsList(), "recipients"));
        }
        if (request.getSendNotification()) {
            body.put("sendNotification", true);
        }
        return graph.execute(new GraphClient.Call("POST",
                resource(request.getTarget(), false) + "/createLink", Map.of(), body));
    }

    /**
     * Grants people or groups access to a drive item.
     *
     * @param graph the Graph API
     * @param request the item, the recipients and the roles
     * @return Graph's answer, a collection of permissions
     */
    static JsonNode invite(GraphClient graph, InviteRecipientsRequest request)
            throws IOException, InterruptedException {
        if (request.getRecipientsCount() == 0) {
            throw new IllegalArgumentException("recipients is required");
        }
        ObjectNode body = GraphClient.object();
        body.set("recipients", recipients(request.getRecipientsList(), "recipients"));
        body.set("roles", roles(request.getRolesList(), true));
        body.put("requireSignIn", request.getRequireSignIn());
        body.put("sendInvitation", request.getSendInvitation());
        if (!request.getMessage().isEmpty()) {
            body.put("message", request.getMessage());
        }
        if (request.hasExpiresAt()) {
            body.put("expirationDateTime", instant(request.getExpiresAt()));
        }
        if (!request.getPassword().isEmpty()) {
            body.put("password", request.getPassword());
        }
        if (request.hasRetainInheritedPermissions()) {
            body.put("retainInheritedPermissions", request.getRetainInheritedPermissions());
        }
        return graph.post(item(request.getItem()) + "/invite", body);
    }

    /**
     * Changes a permission.
     *
     * @param graph the Graph API
     * @param request the permission and its new values
     * @return the permission as Graph answers it
     */
    static JsonNode update(GraphClient graph, UpdatePermissionRequest request)
            throws IOException, InterruptedException {
        if (request.hasExpiresAt() && request.getClearExpiresAt()) {
            throw new IllegalArgumentException("set expires_at or clear_expires_at, not both");
        }
        ObjectNode body = GraphClient.object();
        if (request.getRolesCount() > 0) {
            body.set("roles", roles(request.getRolesList(), true));
        }
        if (request.hasExpiresAt()) {
            body.put("expirationDateTime", instant(request.getExpiresAt()));
        } else if (request.getClearExpiresAt()) {
            body.putNull("expirationDateTime");
        }
        if (request.hasPassword()) {
            body.put("password", request.getPassword());
        }
        if (body.isEmpty()) {
            throw new IllegalArgumentException("nothing to change: set at least one field");
        }
        return graph.patch(permission(owned(request.getTarget()), request.getPermissionId()), body);
    }

    /**
     * Removes a permission.
     *
     * @param graph the Graph API
     * @param target the drive item or list item the permission is on
     * @param permissionId id of the permission
     */
    static void delete(GraphClient graph, PermissionTarget target, String permissionId)
            throws IOException, InterruptedException {
        graph.delete(permission(owned(target), permissionId));
    }

    /**
     * Lets people or groups use a sharing link.
     *
     * @param graph the Graph API
     * @param request the link, the recipients and the roles
     * @return Graph's answer, a collection of permissions
     */
    static JsonNode grant(GraphClient graph, GrantLinkAccessRequest request)
            throws IOException, InterruptedException {
        if (request.getRecipientsCount() == 0) {
            throw new IllegalArgumentException("recipients is required");
        }
        ObjectNode body = GraphClient.object();
        body.set("recipients", recipients(request.getRecipientsList(), "recipients"));
        body.set("roles", roles(request.getRolesList(), true));
        return graph.post(link(request.getItem(), request.getPermissionId()) + "/grant", body);
    }

    /**
     * Takes people or groups off a sharing link.
     *
     * @param graph the Graph API
     * @param request the link and the grantees to remove
     * @return the link's permission as Graph answers it
     */
    static JsonNode revoke(GraphClient graph, RevokeLinkAccessRequest request)
            throws IOException, InterruptedException {
        if (request.getGranteesCount() == 0) {
            throw new IllegalArgumentException("grantees is required");
        }
        ObjectNode body = GraphClient.object();
        body.set("grantees", recipients(request.getGranteesList(), "grantees"));
        return graph.post(link(request.getItem(), request.getPermissionId()) + "/revokeGrants", body);
    }

    /**
     * The target of a permission on a drive item.
     *
     * @param item the drive item
     * @return the target
     */
    static PermissionTarget target(DriveItemRef item) {
        return PermissionTarget.newBuilder().setDriveItem(item).build();
    }

    private static String link(DriveItemRef item, String permissionId) {
        return item(item) + "/permissions/" + GraphPaths.segment(permissionId, "permission_id");
    }

    private static String item(DriveItemRef item) {
        return GraphPaths.item(item.getDriveId(), item.getItemId());
    }

    /** Refuses a list or a site: only an item has permissions that can be changed here. */
    private static PermissionTarget owned(PermissionTarget target) {
        resource(target, false);
        return target;
    }

    private static String resource(PermissionTarget target, boolean containers) {
        return switch (target.getTargetCase()) {
            case DRIVE_ITEM -> item(target.getDriveItem());
            case LIST_ITEM -> GraphPaths.listItem(target.getListItem().getSiteId(),
                    target.getListItem().getListId(), target.getListItem().getItemId());
            case LIST -> {
                if (!containers) {
                    throw new IllegalArgumentException("target must be a drive item or a list item");
                }
                yield GraphPaths.list(target.getList().getSiteId(), target.getList().getListId());
            }
            case SITE_ID -> {
                if (!containers) {
                    throw new IllegalArgumentException("target must be a drive item or a list item");
                }
                yield "/sites/" + GraphPaths.site(target.getSiteId());
            }
            case TARGET_NOT_SET -> throw new IllegalArgumentException("target is required");
        };
    }

    private static ArrayNode recipients(List<SharingRecipient> recipients, String field) {
        ArrayNode array = GraphClient.object().arrayNode();
        for (SharingRecipient recipient : recipients) {
            ObjectNode node = array.addObject();
            switch (recipient.getRecipientCase()) {
                case EMAIL -> node.put("email", filled(recipient.getEmail(), field + ".email"));
                case OBJECT_ID -> node.put("objectId", filled(recipient.getObjectId(), field + ".object_id"));
                case ALIAS -> node.put("alias", filled(recipient.getAlias(), field + ".alias"));
                case RECIPIENT_NOT_SET -> throw new IllegalArgumentException(
                        field + ": each entry needs an email, an object_id or an alias");
            }
        }
        return array;
    }

    private static ArrayNode roles(List<String> roles, boolean required) {
        if (required && roles.isEmpty()) {
            throw new IllegalArgumentException("roles is required");
        }
        ArrayNode array = GraphClient.object().arrayNode();
        for (String role : roles) {
            array.add(filled(role, "roles"));
        }
        return array;
    }

    private static String filled(String value, String field) {
        if (value.isBlank()) {
            throw new IllegalArgumentException(field + " must not be blank");
        }
        return value.trim();
    }

    private static String instant(Timestamp timestamp) {
        return Instant.ofEpochSecond(timestamp.getSeconds(), timestamp.getNanos()).toString();
    }
}
