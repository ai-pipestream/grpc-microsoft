package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import org.junit.jupiter.api.Test;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Clock;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** The receiver's protocol duties: handshake echo, clientState check, 202. */
class NotificationReceiverTest {

    private final HttpClient http = HttpClient.newHttpClient();

    @Test
    void echoesTheDecodedValidationToken() throws Exception {
        try (NotificationReceiver receiver = NotificationReceiver.start(0, "state",
                new NotificationBus(), Clock.systemUTC())) {
            HttpResponse<String> answer = http.send(HttpRequest.newBuilder(URI.create(
                    "http://127.0.0.1:" + receiver.port() + NotificationReceiver.NOTIFICATIONS_PATH
                            + "?validationToken=Validation%3A%20Testing%20client%20application"))
                    .POST(HttpRequest.BodyPublishers.noBody()).build(),
                    HttpResponse.BodyHandlers.ofString());
            assertThat(answer.statusCode()).isEqualTo(200);
            assertThat(answer.body()).isEqualTo("Validation: Testing client application");
            assertThat(answer.headers().firstValue("content-type")).hasValueSatisfying(
                    type -> assertThat(type).startsWith("text/plain"));
        }
    }

    @Test
    void publishesOnlyItemsWithTheRightClientState() throws Exception {
        NotificationBus bus = new NotificationBus();
        try (NotificationReceiver receiver = NotificationReceiver.start(0, "state", bus,
                Clock.systemUTC()); NotificationBus.Listener listener = bus.listen()) {
            String body = "{\"value\":[{\"subscriptionId\":\"a\",\"clientState\":\"state\","
                    + "\"changeType\":\"updated\",\"resource\":\"drives/x/root\"},"
                    + "{\"subscriptionId\":\"b\",\"clientState\":\"wrong\",\"changeType\":\"updated\"}]}";
            HttpResponse<String> answer = http.send(HttpRequest.newBuilder(URI.create(
                    "http://127.0.0.1:" + receiver.port() + NotificationReceiver.NOTIFICATIONS_PATH))
                    .POST(HttpRequest.BodyPublishers.ofString(body)).build(),
                    HttpResponse.BodyHandlers.ofString());
            assertThat(answer.statusCode()).isEqualTo(202);
            ChangeNotification first = listener.queue().poll(5, TimeUnit.SECONDS);
            assertThat(first).isNotNull();
            assertThat(first.getSubscriptionId()).isEqualTo("a");
            assertThat(listener.queue().poll(200, TimeUnit.MILLISECONDS)).isNull();
        }
    }

    @Test
    void refusesAMissingOrOverlongClientState() {
        assertThatThrownBy(() -> NotificationReceiver.start(0, "", new NotificationBus(),
                Clock.systemUTC())).isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> NotificationReceiver.start(0, "x".repeat(129),
                new NotificationBus(), Clock.systemUTC())).isInstanceOf(IllegalArgumentException.class);
    }
}
