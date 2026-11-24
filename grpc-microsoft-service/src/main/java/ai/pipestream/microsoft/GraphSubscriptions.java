package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.CreateSubscriptionRequest;
import ai.pipestream.microsoft.v1.Subscription;
import ai.pipestream.microsoft.v1.SubscriptionChangeType;
import ai.pipestream.microsoft.v1.SubscriptionResource;
import ai.pipestream.microsoft.v1.UpdateSubscriptionRequest;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;
import io.grpc.Status;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.time.Instant;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/** Graph {@code /subscriptions} calls and their mapping to the typed contract. */
final class GraphSubscriptions {

    /** Longest lifetime Graph allows for drive root and list subscriptions. */
    static final Duration DRIVE_AND_LIST_MAX = Duration.ofMinutes(42_300);
    /** Longest lifetime Graph allows for directory (group, user) subscriptions. */
    static final Duration DIRECTORY_MAX = Duration.ofMinutes(41_760);
    /** Margin kept below the maximum so clock skew does not push a request over it. */
    static final Duration MARGIN = Duration.ofMinutes(5);

    private static final ObjectMapper JSON = new ObjectMapper();

    private final GraphClient graph;

    GraphSubscriptions(GraphClient graph) {
        this.graph = graph;
    }

    /** Graph resource path of a typed resource. */
    static String path(SubscriptionResource resource) {
        return switch (resource.getTargetCase()) {
            case DRIVE_ROOT -> "/drives/" + segment(resource.getDriveRoot()) + "/root";
            case LIST -> "/sites/" + segment(resource.getList().getSiteId()) + "/lists/"
                    + segment(resource.getList().getListId());
            case GROUP_MEMBERS -> "/groups/" + segment(resource.getGroupMembers()) + "/members";
            case GROUP -> "/groups/" + segment(resource.getGroup());
            case USER -> "/users/" + segment(resource.getUser());
            case PATH -> resource.getPath();
            case TARGET_NOT_SET -> throw Status.INVALID_ARGUMENT
                    .withDescription("resource is required").asRuntimeException();
        };
    }

    /** Longest lifetime for a resource path, minus {@link #MARGIN}. */
    static Duration maxLifetime(String resourcePath) {
        String path = resourcePath.toLowerCase(Locale.ROOT);
        boolean directory = path.startsWith("/groups") || path.startsWith("groups")
                || path.startsWith("/users") || path.startsWith("users");
        return (directory ? DIRECTORY_MAX : DRIVE_AND_LIST_MAX).minus(MARGIN);
    }

    /** POST /subscriptions. */
    Subscription create(CreateSubscriptionRequest request, String clientState, Instant now)
            throws IOException, InterruptedException {
        if (request.getChangeTypesCount() == 0) {
            throw Status.INVALID_ARGUMENT.withDescription("change_types is required")
                    .asRuntimeException();
        }
        if (request.getNotificationUrl().isBlank()) {
            throw Status.INVALID_ARGUMENT.withDescription("notification_url is required")
                    .asRuntimeException();
        }
        String resource = path(request.getResource());
        ObjectNode body = JSON.createObjectNode()
                .put("changeType", changeTypes(request.getChangeTypesList()))
                .put("notificationUrl", request.getNotificationUrl())
                .put("resource", resource)
                .put("expirationDateTime", expiration(request.hasExpirationDateTime()
                        ? request.getExpirationDateTime() : null, resource, now).toString())
                .put("clientState", clientState);
        if (!request.getLifecycleNotificationUrl().isBlank()) {
            body.put("lifecycleNotificationUrl", request.getLifecycleNotificationUrl());
        }
        Map<String, String> headers = request.getIncludeSecurityWebhooks()
                ? Map.of("Prefer", "includesecuritywebhooks") : Map.of();
        return subscription(graph.execute(new GraphClient.Call("POST", "/subscriptions", headers,
                body)).body());
    }

    /** GET /subscriptions/{id}. */
    Subscription get(String id) throws IOException, InterruptedException {
        return subscription(graph.get("/subscriptions/" + segment(required(id))));
    }

    /** GET /subscriptions, following {@code @odata.nextLink}. */
    List<Subscription> list() throws IOException, InterruptedException {
        List<Subscription> out = new ArrayList<>();
        String next = "/subscriptions";
        while (next != null && !next.isBlank()) {
            JsonNode page = graph.get(next);
            page.path("value").forEach(node -> out.add(subscription(node)));
            next = page.path("@odata.nextLink").asText("");
        }
        return out;
    }

    /** PATCH /subscriptions/{id}. */
    Subscription update(UpdateSubscriptionRequest request, Instant now)
            throws IOException, InterruptedException {
        String id = required(request.getSubscriptionId());
        String resource = request.hasExpirationDateTime() ? "" : get(id).getResource();
        ObjectNode body = JSON.createObjectNode().put("expirationDateTime", expiration(
                request.hasExpirationDateTime() ? request.getExpirationDateTime() : null, resource,
                now).toString());
        if (!request.getNotificationUrl().isBlank()) {
            body.put("notificationUrl", request.getNotificationUrl());
        }
        return subscription(graph.patch("/subscriptions/" + segment(id), body));
    }

    /** POST /subscriptions/{id}/reauthorize, with an empty JSON body so a length is sent. */
    void reauthorize(String id) throws IOException, InterruptedException {
        graph.execute(new GraphClient.Call("POST", "/subscriptions/" + segment(required(id))
                + "/reauthorize", Map.of(), JSON.createObjectNode()));
    }

    /** DELETE /subscriptions/{id}. */
    void delete(String id) throws IOException, InterruptedException {
        graph.delete("/subscriptions/" + segment(required(id)));
    }

    private static Instant expiration(Timestamp requested, String resource, Instant now) {
        if (requested != null) {
            return Instant.ofEpochSecond(requested.getSeconds(), requested.getNanos());
        }
        return now.plus(maxLifetime(resource));
    }

    static String changeTypes(List<SubscriptionChangeType> types) {
        List<String> names = new ArrayList<>();
        for (SubscriptionChangeType type : types) {
            String name = switch (type) {
                case SUBSCRIPTION_CHANGE_TYPE_CREATED -> "created";
                case SUBSCRIPTION_CHANGE_TYPE_UPDATED -> "updated";
                case SUBSCRIPTION_CHANGE_TYPE_DELETED -> "deleted";
                default -> throw Status.INVALID_ARGUMENT
                        .withDescription("unsupported change type " + type).asRuntimeException();
            };
            if (!names.contains(name)) {
                names.add(name);
            }
        }
        return String.join(",", names);
    }

    static SubscriptionChangeType changeType(String name) {
        return switch (name.trim().toLowerCase(Locale.ROOT)) {
            case "created" -> SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_CREATED;
            case "updated" -> SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED;
            case "deleted" -> SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_DELETED;
            default -> SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UNSPECIFIED;
        };
    }

    static Subscription subscription(JsonNode node) {
        Subscription.Builder out = Subscription.newBuilder()
                .setId(node.path("id").asText(""))
                .setResource(node.path("resource").asText(""))
                .setNotificationUrl(node.path("notificationUrl").asText(""))
                .setLifecycleNotificationUrl(node.path("lifecycleNotificationUrl").asText(""))
                .setApplicationId(node.path("applicationId").asText(""))
                .setCreatorId(node.path("creatorId").asText(""))
                .setIncludeResourceData(node.path("includeResourceData").asBoolean(false))
                .setLatestSupportedTlsVersion(node.path("latestSupportedTlsVersion").asText(""))
                .setNotificationUrlAppId(node.path("notificationUrlAppId").asText(""));
        for (String type : node.path("changeType").asText("").split(",", -1)) {
            SubscriptionChangeType parsed = changeType(type);
            if (parsed != SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UNSPECIFIED) {
                out.addChangeTypes(parsed);
            }
        }
        Timestamp expires = timestamp(node.path("expirationDateTime").asText(""));
        if (expires != null) {
            out.setExpirationDateTime(expires);
        }
        return out.build();
    }

    /** Parses a Graph date-time; {@code null} when absent or unreadable. */
    static Timestamp timestamp(String text) {
        if (text == null || text.isBlank()) {
            return null;
        }
        try {
            Instant instant = Instant.parse(text);
            return Timestamp.newBuilder().setSeconds(instant.getEpochSecond())
                    .setNanos(instant.getNano()).build();
        } catch (DateTimeParseException e) {
            return null;
        }
    }

    private static String required(String id) {
        if (id == null || id.isBlank()) {
            throw Status.INVALID_ARGUMENT.withDescription("subscription_id is required")
                    .asRuntimeException();
        }
        return id;
    }

    private static String segment(String value) {
        return URLEncoder.encode(value, StandardCharsets.UTF_8).replace("+", "%20")
                .replace("%2C", ",");
    }
}
