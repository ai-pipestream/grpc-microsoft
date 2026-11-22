package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.protobuf.Timestamp;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpServer;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Clock;
import java.time.Instant;
import java.util.Locale;
import java.util.concurrent.Executors;
import java.util.function.Consumer;

/**
 * HTTPS-facing receiver for Graph change and lifecycle notifications. Answers
 * the validation handshake, drops items whose clientState does not match,
 * publishes the rest to a {@link NotificationBus}, and answers 202. It does no
 * other work in the request, so Graph's 3 second budget holds.
 */
final class NotificationReceiver implements AutoCloseable {

    /** Environment variable for the receiver's listen port; unset = receiver off. */
    static final String ENV_PORT = "MICROSOFT_NOTIFICATION_PORT";
    /** Environment variable for the clientState secret sent with every subscription. */
    static final String ENV_CLIENT_STATE = "MICROSOFT_NOTIFICATION_CLIENT_STATE";
    /** Path of change notifications. */
    static final String NOTIFICATIONS_PATH = "/graph/notifications";
    /** Path of lifecycle notifications. */
    static final String LIFECYCLE_PATH = "/graph/lifecycle";
    /** Longest notification body read, in bytes. */
    static final int MAX_BODY_BYTES = 4 * 1024 * 1024;

    private static final System.Logger LOG = System.getLogger(NotificationReceiver.class.getName());
    private static final ObjectMapper JSON = new ObjectMapper();

    private final HttpServer server;
    private final byte[] clientState;
    private final NotificationBus bus;
    private final Clock clock;
    private volatile Consumer<JsonNode> tap = envelope -> { };

    private NotificationReceiver(HttpServer server, String clientState, NotificationBus bus,
            Clock clock) {
        this.server = server;
        this.clientState = clientState.getBytes(StandardCharsets.UTF_8);
        this.bus = bus;
        this.clock = clock;
    }

    /**
     * Starts a receiver on {@code port} (0 = any free port).
     *
     * @param port listen port
     * @param clientState the secret every subscription carries
     * @param bus where accepted notifications go
     * @param clock source of received-at times
     * @return the started receiver
     * @throws IOException if the port cannot be bound
     */
    static NotificationReceiver start(int port, String clientState, NotificationBus bus, Clock clock)
            throws IOException {
        if (clientState == null || clientState.isBlank()) {
            throw new IllegalArgumentException(ENV_CLIENT_STATE + " is required for the receiver");
        }
        if (clientState.length() > 128) {
            throw new IllegalArgumentException(ENV_CLIENT_STATE + " is longer than 128 characters");
        }
        HttpServer server = HttpServer.create(new InetSocketAddress(port), 0);
        NotificationReceiver receiver = new NotificationReceiver(server, clientState, bus, clock);
        server.createContext(NOTIFICATIONS_PATH, receiver::handle);
        server.createContext(LIFECYCLE_PATH, receiver::handle);
        server.setExecutor(Executors.newVirtualThreadPerTaskExecutor());
        server.start();
        return receiver;
    }

    /** Sets a consumer that sees every notification body before it is checked. */
    void tap(Consumer<JsonNode> consumer) {
        this.tap = consumer;
    }

    /** The bound port. */
    int port() {
        return server.getAddress().getPort();
    }

    @Override
    public void close() {
        server.stop(0);
    }

    private void handle(HttpExchange exchange) throws IOException {
        try (exchange) {
            String token = query(exchange.getRequestURI().getRawQuery(), "validationToken");
            if (token != null) {
                byte[] body = token.getBytes(StandardCharsets.UTF_8);
                exchange.getResponseHeaders().set("Content-Type", "text/plain; charset=utf-8");
                exchange.sendResponseHeaders(200, body.length);
                try (OutputStream out = exchange.getResponseBody()) {
                    out.write(body);
                }
                return;
            }
            if (!"POST".equalsIgnoreCase(exchange.getRequestMethod())) {
                exchange.sendResponseHeaders(405, -1);
                return;
            }
            byte[] raw;
            try (InputStream in = exchange.getRequestBody()) {
                raw = in.readNBytes(MAX_BODY_BYTES + 1);
            }
            if (raw.length > MAX_BODY_BYTES) {
                exchange.sendResponseHeaders(413, -1);
                return;
            }
            int accepted;
            try {
                accepted = accept(JSON.readTree(raw));
            } catch (IOException e) {
                exchange.sendResponseHeaders(400, -1);
                return;
            }
            LOG.log(System.Logger.Level.DEBUG, "accepted {0} notifications", accepted);
            exchange.sendResponseHeaders(202, -1);
        }
    }

    /** Publishes every item with a matching clientState; answers how many. */
    int accept(JsonNode envelope) {
        tap.accept(envelope);
        int accepted = 0;
        for (JsonNode item : envelope.path("value")) {
            byte[] state = item.path("clientState").asText("").getBytes(StandardCharsets.UTF_8);
            if (!MessageDigest.isEqual(state, clientState)) {
                LOG.log(System.Logger.Level.WARNING, "dropped a notification with a wrong clientState");
                continue;
            }
            bus.publish(notification(item, clock.instant()));
            accepted++;
        }
        return accepted;
    }

    static ChangeNotification notification(JsonNode item, Instant receivedAt) {
        ChangeNotification.Builder out = ChangeNotification.newBuilder()
                .setSubscriptionId(item.path("subscriptionId").asText(""))
                .setChangeType(GraphSubscriptions.changeType(item.path("changeType").asText("")))
                .setResource(item.path("resource").asText(""))
                .setTenantId(item.path("tenantId").asText(""))
                .setLifecycleEvent(lifecycle(item.path("lifecycleEvent").asText("")))
                .setResourceId(item.path("resourceData").path("id").asText(""))
                .setReceivedAt(Timestamp.newBuilder().setSeconds(receivedAt.getEpochSecond())
                        .setNanos(receivedAt.getNano()));
        Timestamp expires = GraphSubscriptions.timestamp(
                item.path("subscriptionExpirationDateTime").asText(""));
        if (expires != null) {
            out.setSubscriptionExpirationDateTime(expires);
        }
        return out.build();
    }

    static LifecycleEvent lifecycle(String name) {
        return switch (name.toLowerCase(Locale.ROOT)) {
            case "reauthorizationrequired" -> LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED;
            case "subscriptionremoved" -> LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED;
            case "missed" -> LifecycleEvent.LIFECYCLE_EVENT_MISSED;
            default -> LifecycleEvent.LIFECYCLE_EVENT_UNSPECIFIED;
        };
    }

    /** The decoded value of one query parameter; {@code null} when absent. */
    static String query(String rawQuery, String name) {
        if (rawQuery == null || rawQuery.isEmpty()) {
            return null;
        }
        for (String pair : rawQuery.split("&", -1)) {
            int eq = pair.indexOf('=');
            String key = eq < 0 ? pair : pair.substring(0, eq);
            if (URLDecoder.decode(key, StandardCharsets.UTF_8).equals(name)) {
                return eq < 0 ? "" : URLDecoder.decode(pair.substring(eq + 1), StandardCharsets.UTF_8);
            }
        }
        return null;
    }
}
