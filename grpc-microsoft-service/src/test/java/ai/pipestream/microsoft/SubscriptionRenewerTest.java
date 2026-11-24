package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import ai.pipestream.microsoft.v1.RenewalAction;
import ai.pipestream.microsoft.v1.Subscription;
import ai.pipestream.microsoft.v1.SubscriptionRenewal;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Clock;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The renewer's decisions that a live run cannot force: when a subscription is
 * due, repeated lifecycle notifications, failures and a removal Graph
 * announces. The subscription bodies come from the recorded
 * {@code CreateSubscription} answer.
 */
class SubscriptionRenewerTest {

    private static final String CLIENT_STATE = "renewer-test-state";
    private static final String BASE = "https://graph.microsoft.com/v1.0";

    /** A clock the test moves. */
    private static final class TestClock extends Clock {
        private Instant now;

        private TestClock(Instant now) {
            this.now = now;
        }

        @Override
        public ZoneId getZone() {
            return ZoneOffset.UTC;
        }

        @Override
        public Clock withZone(ZoneId zone) {
            return this;
        }

        @Override
        public Instant instant() {
            return now;
        }
    }

    /** Answers subscription calls from memory and remembers each request. */
    private final class FakeSubscriptions implements GraphTransport {
        private final List<String> requests = new ArrayList<>();
        private final List<String> preferences = new ArrayList<>();
        private int failuresLeft;
        private boolean gone;
        private int created;

        @Override
        public <T> HttpResponse<T> send(HttpRequest request, HttpResponse.BodyHandler<T> handler)
                throws IOException {
            String route = request.method() + " "
                    + request.uri().toString().substring(BASE.length());
            requests.add(route);
            JsonNode sent = Recordings.JSON.readTree(Recordings.body(request));
            if (failuresLeft > 0) {
                failuresLeft--;
                return answer(request, 403, error("accessDenied"), handler);
            }
            if (request.method().equals("PATCH")) {
                if (gone) {
                    return answer(request, 404, error("ResourceNotFound"), handler);
                }
                ObjectNode body = template.deepCopy();
                body.put("id", request.uri().getPath().substring(
                        request.uri().getPath().lastIndexOf('/') + 1));
                body.put("expirationDateTime", sent.path("expirationDateTime").asText());
                return answer(request, 200, body, handler);
            }
            if (request.method().equals("POST")) {
                preferences.add(request.headers().firstValue("Prefer").orElse(""));
                gone = false;
                ObjectNode body = template.deepCopy();
                body.put("id", "created-" + ++created);
                body.put("resource", sent.path("resource").asText());
                body.put("changeType", sent.path("changeType").asText());
                body.put("notificationUrl", sent.path("notificationUrl").asText());
                body.put("lifecycleNotificationUrl", sent.path("lifecycleNotificationUrl").asText());
                body.put("expirationDateTime", sent.path("expirationDateTime").asText());
                body.put("clientState", sent.path("clientState").asText());
                return answer(request, 201, body, handler);
            }
            throw new IOException("unexpected request " + route);
        }

        private <T> HttpResponse<T> answer(HttpRequest request, int status, JsonNode body,
                HttpResponse.BodyHandler<T> handler) {
            return Recordings.respond(request, status,
                    Recordings.headers(Recordings.JSON.createObjectNode()
                            .put("content-type", "application/json")),
                    body.toString().getBytes(StandardCharsets.UTF_8), handler);
        }

        private JsonNode error(String code) {
            ObjectNode out = Recordings.JSON.createObjectNode();
            out.putObject("error").put("code", code).put("message", code);
            return out;
        }
    }

    private final TestClock clock = new TestClock(Instant.parse("2026-11-24T00:00:00Z"));
    private final NotificationBus bus = new NotificationBus();
    private final FakeSubscriptions graph = new FakeSubscriptions();
    private ObjectNode template;
    private SubscriptionRenewer renewer;

    @BeforeEach
    void setUp() throws IOException {
        template = (ObjectNode) Recordings.JSON.readTree(ChangeRecordings.dir()
                .resolve("01-create-subscription/001-POST-subscriptions.json").toFile())
                .path("response").path("body").path("json");
        GraphSubscriptions calls = new GraphSubscriptions(new GraphClient(BASE, () -> "token",
                AuthMode.AUTH_MODE_APPLICATION, graph));
        renewer = new SubscriptionRenewer(connection -> calls, bus, CLIENT_STATE, clock,
                Duration.ofHours(24));
    }

    private Subscription subscription(String id, Instant expires) {
        ObjectNode body = template.deepCopy();
        body.put("id", id);
        body.put("expirationDateTime", expires.toString());
        return GraphSubscriptions.subscription(body);
    }

    private SubscriptionRenewal only() {
        assertThat(renewer.renewals()).hasSize(1);
        return renewer.renewals().get(0);
    }

    @Test
    void aSubscriptionIsRenewedOnlyOnceItsExpirationIsNear() {
        renewer.created("", true, subscription("sub-1", clock.instant().plus(Duration.ofDays(3))));
        renewer.runOnce();
        assertThat(graph.requests).isEmpty();
        assertThat(only().getRenewAt().getSeconds())
                .isEqualTo(clock.instant().plus(Duration.ofDays(2)).getEpochSecond());

        clock.now = clock.now.plus(Duration.ofDays(2));
        renewer.runOnce();
        assertThat(graph.requests).containsExactly("PATCH /subscriptions/sub-1");
        assertThat(only().getLastAction()).isEqualTo(RenewalAction.RENEWAL_ACTION_RENEWED);
        assertThat(only().getSubscription().getExpirationDateTime().getSeconds()).isEqualTo(
                clock.instant().plus(GraphSubscriptions.maxLifetime("/drives/x/root")).getEpochSecond());

        renewer.runOnce();
        assertThat(graph.requests).hasSize(1);
    }

    @Test
    void aRepeatedReauthorizationRequestIsAnsweredOnce() {
        renewer.created("", true, subscription("sub-1", clock.instant().plus(Duration.ofDays(20))));
        ChangeNotification asked = ChangeNotification.newBuilder().setSubscriptionId("sub-1")
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED).build();
        renewer.lifecycle(asked);
        clock.now = clock.now.plus(Duration.ofMinutes(9));
        renewer.lifecycle(asked);
        assertThat(graph.requests).containsExactly("PATCH /subscriptions/sub-1");
        assertThat(only().getLastAction()).isEqualTo(RenewalAction.RENEWAL_ACTION_REAUTHORIZED);

        clock.now = clock.now.plus(Duration.ofMinutes(2));
        renewer.lifecycle(asked);
        assertThat(graph.requests).hasSize(2);
    }

    @Test
    void aFailedRenewalIsReportedAndTriedAgainOnTheNextCheck() {
        renewer.created("", true, subscription("sub-1", clock.instant().plus(Duration.ofDays(20))));
        graph.failuresLeft = 1;
        renewer.lifecycle(ChangeNotification.newBuilder().setSubscriptionId("sub-1")
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED).build());
        assertThat(only().getLastError()).startsWith("PERMISSION_DENIED");
        assertThat(only().getLastAction()).isEqualTo(RenewalAction.RENEWAL_ACTION_UNSPECIFIED);

        renewer.runOnce();
        assertThat(graph.requests).hasSize(2);
        assertThat(only().getLastError()).isEmpty();
        assertThat(only().getLastAction()).isEqualTo(RenewalAction.RENEWAL_ACTION_REAUTHORIZED);
    }

    @Test
    void aSubscriptionGraphRemovedIsCreatedAgainAndWatchersAreTold() throws Exception {
        renewer.created("", false, subscription("sub-1", clock.instant().plus(Duration.ofDays(20))));
        try (NotificationBus.Listener listener = bus.listen()) {
            renewer.lifecycle(ChangeNotification.newBuilder().setSubscriptionId("sub-1")
                    .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED).build());
            assertThat(graph.requests).containsExactly("POST /subscriptions");
            assertThat(graph.preferences).containsExactly("");
            SubscriptionRenewal renewal = only();
            assertThat(renewal.getSubscription().getId()).isEqualTo("created-1");
            assertThat(renewal.getReplacedSubscriptionId()).isEqualTo("sub-1");
            assertThat(renewal.getLastAction()).isEqualTo(RenewalAction.RENEWAL_ACTION_RECREATED);
            ChangeNotification missed = listener.queue().poll(5, TimeUnit.SECONDS);
            assertThat(missed).isNotNull();
            assertThat(missed.getSubscriptionId()).isEqualTo("created-1");
            assertThat(missed.getLifecycleEvent()).isEqualTo(LifecycleEvent.LIFECYCLE_EVENT_MISSED);
        }
    }

    @Test
    void aFailedCreationIsTriedAgainOnTheNextCheck() {
        renewer.created("", true, subscription("sub-1", clock.instant().plus(Duration.ofDays(20))));
        graph.failuresLeft = 1;
        renewer.lifecycle(ChangeNotification.newBuilder().setSubscriptionId("sub-1")
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED).build());
        assertThat(only().getSubscription().getId()).isEqualTo("sub-1");
        assertThat(only().getLastError()).startsWith("PERMISSION_DENIED");

        renewer.runOnce();
        assertThat(graph.requests).containsExactly("POST /subscriptions", "POST /subscriptions");
        assertThat(only().getSubscription().getId()).isEqualTo("created-1");
        assertThat(graph.preferences).containsExactly("includesecuritywebhooks");
    }

    @Test
    void onlySubscriptionsWithThisServersClientStateAreAdopted() {
        Subscription mine = subscription("sub-mine", clock.instant().plus(Duration.ofDays(20)));
        Subscription other = subscription("sub-other", clock.instant().plus(Duration.ofDays(20)));
        assertThat(renewer.seen("", new GraphSubscriptions.Entry(other, "another-state"))).isFalse();
        assertThat(renewer.seen("c-1", new GraphSubscriptions.Entry(mine, CLIENT_STATE))).isTrue();
        assertThat(renewer.seen("c-1", new GraphSubscriptions.Entry(mine, CLIENT_STATE))).isFalse();
        assertThat(only().getSubscription().getId()).isEqualTo("sub-mine");
        assertThat(only().getConnectionId()).isEqualTo("c-1");
    }

    @Test
    void anAdoptedDriveSubscriptionIsCreatedAgainWithSecurityWebhooks() {
        Subscription adopted = subscription("sub-1", clock.instant().plus(Duration.ofHours(1)));
        renewer.seen("", new GraphSubscriptions.Entry(adopted, CLIENT_STATE));
        graph.gone = true;
        renewer.runOnce();
        assertThat(graph.requests).containsExactly("PATCH /subscriptions/sub-1", "POST /subscriptions");
        assertThat(graph.preferences).containsExactly("includesecuritywebhooks");
        assertThat(only().getReplacedSubscriptionId()).isEqualTo("sub-1");
    }

    @Test
    void aDeletedSubscriptionIsNoLongerKept() {
        renewer.created("", true, subscription("sub-1", clock.instant()));
        renewer.deleted("sub-1");
        renewer.runOnce();
        assertThat(graph.requests).isEmpty();
        assertThat(renewer.renewals()).isEmpty();
    }

    @Test
    void theEnvironmentTurnsRenewalOffAndSetsTheMargin() {
        assertThat(SubscriptionRenewer.enabled(null)).isTrue();
        assertThat(SubscriptionRenewer.enabled(" OFF ")).isFalse();
        assertThat(SubscriptionRenewer.renewBefore(null)).isEqualTo(Duration.ofHours(24));
        assertThat(SubscriptionRenewer.renewBefore("90")).isEqualTo(Duration.ofMinutes(90));
        assertThat(SubscriptionRenewer.renewBefore("soon")).isEqualTo(Duration.ofHours(24));
    }
}
