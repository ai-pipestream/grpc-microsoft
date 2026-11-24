package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import com.fasterxml.jackson.databind.JsonNode;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.TestFactory;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Clock;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Runs {@link ChangeScenarios} against what {@link MicrosoftLiveChangeIT}
 * recorded. Graph exchanges replay through {@link ReplayTransport}; the
 * recorded notification bodies are posted to a local receiver over HTTP, so
 * the receiver, the WatchNotifications stream and the notification-started
 * WatchChanges pass all run for real. Steps must run in order.
 */
class MicrosoftChangeReplayTest {

    private ExecutorService executor;
    private Server server;
    private ManagedChannel channel;
    private NotificationReceiver receiver;
    private ChangeScenarios scenarios;

    @AfterEach
    void stop() {
        if (scenarios != null) {
            scenarios.close();
        }
        if (receiver != null) {
            receiver.close();
        }
        if (channel != null) {
            channel.shutdownNow();
        }
        if (server != null) {
            server.shutdownNow();
        }
        if (executor != null) {
            executor.shutdownNow();
        }
    }

    @TestFactory
    List<DynamicTest> steps() throws Exception {
        Path recordings = ChangeRecordings.dir();
        Path runFile = recordings.resolve(Recordings.RUN_FILE);
        boolean recorded = Files.isRegularFile(runFile);
        List<DynamicTest> tests = new ArrayList<>();
        if (!recorded) {
            tests.add(DynamicTest.dynamicTest("recordings", () -> assumeTrue(false,
                    "no change-feed recording yet under " + recordings
                            + "; record one with liveChangesTest")));
            return tests;
        }
        JsonNode run = Recordings.JSON.readTree(runFile.toFile());
        String suffix = run.path("suffix").asText();
        ChangeScenarios.Backend backend = new ChangeScenarios.Backend(run.path("driveId").asText(),
                suffix, Instant.parse(run.path("startedAt").asText()),
                run.path("notificationUrl").asText(), run.path("lifecycleUrl").asText());
        Map<String, byte[]> blobs = new ConcurrentHashMap<>();
        blobs.put(Recordings.sha256(ChangeRecordings.content(suffix)), ChangeRecordings.content(suffix));
        ReplayTransport transport = new ReplayTransport(blobs);
        String graphBaseUrl = run.path("graphBaseUrl").asText();
        GraphClient client = new GraphClient(graphBaseUrl, () -> "replay-token",
                AuthMode.AUTH_MODE_APPLICATION, transport);
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("replay-tenant").clientId("replay-client").clientSecret("replay-secret")
                .graphBaseUrl(graphBaseUrl).authority(graphBaseUrl).build();
        NotificationBus receiverBus = new NotificationBus();
        NotificationBus serviceBus = new NotificationBus();
        receiver = NotificationReceiver.start(0, backend.clientState(), receiverBus,
                Clock.systemUTC());
        MicrosoftChangeGrpcService service = new MicrosoftChangeGrpcService(
                new MicrosoftGrpcService(config, new GraphFiles(client),
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES),
                serviceBus, true, backend.clientState(), Clock.systemUTC());
        executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor).addService(service).build()
                .start();
        channel = InProcessChannelBuilder.forName(name).build();
        HttpClient http = HttpClient.newHttpClient();
        URI post = URI.create("http://127.0.0.1:" + receiver.port()
                + NotificationReceiver.NOTIFICATIONS_PATH);
        ChangeScenarios.Notifications replay = new ChangeScenarios.Notifications() {
            @Override
            public void arrive(String stepId) throws Exception {
                Path file = recordings.resolve(stepId).resolve(ChangeRecordings.NOTIFICATIONS_FILE);
                if (!Files.isRegularFile(file)) {
                    return;
                }
                for (JsonNode body : Recordings.JSON.readTree(file.toFile())) {
                    HttpResponse<String> answer = http.send(HttpRequest.newBuilder(post)
                            .header("Content-Type", "application/json")
                            .POST(HttpRequest.BodyPublishers.ofString(body.toString())).build(),
                            HttpResponse.BodyHandlers.ofString());
                    assertThat(answer.statusCode()).as("receiver answer").isEqualTo(202);
                }
            }

            @Override
            public void settle(Duration quiet) {
                // Nothing arrives on its own during replay.
            }
        };
        scenarios = new ChangeScenarios(MicrosoftChangeServiceGrpc.newBlockingStub(channel),
                MicrosoftChangeServiceGrpc.newStub(channel), client, backend, receiverBus,
                serviceBus, replay);
        NotificationBus.Listener listener = scenarios.listen();
        for (ChangeScenarios.Scenario scenario : scenarios.all()) {
            tests.add(DynamicTest.dynamicTest(scenario.id(), () -> {
                Path directory = recordings.resolve(scenario.id());
                assumeTrue(Files.isDirectory(directory), "no recording of " + scenario.id());
                transport.begin(directory);
                WriteScenarios.Outcome outcome = scenarios.run(scenario);
                List<String> problems = transport.finish();
                assumeTrue(outcome.status() != WriteScenarios.Status.SKIPPED, outcome.detail());
                assertThat(problems).as("requests that differ from the recording").isEmpty();
                assertThat(outcome.status()).as(outcome.detail())
                        .isEqualTo(WriteScenarios.Status.PASS);
            }));
        }
        tests.add(DynamicTest.dynamicTest("close", listener::close));
        return tests;
    }
}
