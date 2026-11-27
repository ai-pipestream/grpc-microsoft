package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.GraphGroup;
import ai.pipestream.microsoft.v1.UpdateGroupRequest;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import java.io.IOException;

/** Creates, changes and deletes Entra groups and their direct members through Graph. */
final class DirectoryWriter {

    private DirectoryWriter() {
    }

    /**
     * Creates a group.
     *
     * @param graph the Graph API
     * @param group the group to create
     * @return the group as Graph created it
     */
    static JsonNode createGroup(GraphClient graph, GraphGroup group) throws IOException, InterruptedException {
        if (group.getDisplayName().isBlank()) {
            throw new IllegalArgumentException("group.display_name is required");
        }
        if (group.getMailNickname().isBlank()) {
            throw new IllegalArgumentException("group.mail_nickname is required");
        }
        if (!group.getId().isEmpty() || !group.getMail().isEmpty()
                || !group.getSecurityIdentifier().isEmpty()) {
            throw new IllegalArgumentException(
                    "group.id, group.mail and group.security_identifier are set by Graph");
        }
        ObjectNode body = GraphClient.object()
                .put("displayName", group.getDisplayName().trim())
                .put("mailNickname", group.getMailNickname().trim())
                .put("mailEnabled", group.hasMailEnabled() && group.getMailEnabled())
                .put("securityEnabled", !group.hasSecurityEnabled() || group.getSecurityEnabled());
        if (!group.getDescription().isEmpty()) {
            body.put("description", group.getDescription());
        }
        if (!group.getVisibility().isEmpty()) {
            body.put("visibility", group.getVisibility());
        }
        ArrayNode types = body.putArray("groupTypes");
        group.getGroupTypesList().forEach(types::add);
        return graph.post("/groups", body);
    }

    /**
     * Changes a group.
     *
     * @param graph the Graph API
     * @param request names the group and the new values
     */
    static void updateGroup(GraphClient graph, UpdateGroupRequest request)
            throws IOException, InterruptedException {
        ObjectNode body = GraphClient.object();
        if (request.hasDisplayName()) {
            if (request.getDisplayName().isBlank()) {
                throw new IllegalArgumentException("display_name must not be blank");
            }
            body.put("displayName", request.getDisplayName().trim());
        }
        if (request.hasDescription()) {
            if (request.getDescription().isEmpty()) {
                body.putNull("description");
            } else {
                body.put("description", request.getDescription());
            }
        }
        if (request.hasMailNickname()) {
            if (request.getMailNickname().isBlank()) {
                throw new IllegalArgumentException("mail_nickname must not be blank");
            }
            body.put("mailNickname", request.getMailNickname().trim());
        }
        if (request.hasVisibility()) {
            if (request.getVisibility().isBlank()) {
                throw new IllegalArgumentException("visibility must not be blank");
            }
            body.put("visibility", request.getVisibility().trim());
        }
        if (body.isEmpty()) {
            throw new IllegalArgumentException("nothing to change: set at least one field");
        }
        graph.patch(group(request.getGroupId()), body);
    }

    /**
     * Deletes a group.
     *
     * @param graph the Graph API
     * @param groupId Graph object id of the group
     */
    static void deleteGroup(GraphClient graph, String groupId) throws IOException, InterruptedException {
        graph.delete(group(groupId));
    }

    /**
     * Adds a directory object to a group.
     *
     * @param graph the Graph API
     * @param groupId Graph object id of the group
     * @param memberId Graph object id of the new member
     */
    static void addMember(GraphClient graph, String groupId, String memberId)
            throws IOException, InterruptedException {
        ObjectNode body = GraphClient.object().put("@odata.id", graph.baseUrl() + "/directoryObjects/"
                + GraphPaths.segment(memberId, "member_id"));
        graph.post(group(groupId) + "/members/$ref", body);
    }

    /**
     * Takes a direct member out of a group.
     *
     * @param graph the Graph API
     * @param groupId Graph object id of the group
     * @param memberId Graph object id of the member
     */
    static void removeMember(GraphClient graph, String groupId, String memberId)
            throws IOException, InterruptedException {
        graph.delete(group(groupId) + "/members/" + GraphPaths.segment(memberId, "member_id") + "/$ref");
    }

    /** {@code /groups/{group-id}}. */
    static String group(String groupId) {
        return "/groups/" + GraphPaths.segment(groupId, "group_id");
    }
}
