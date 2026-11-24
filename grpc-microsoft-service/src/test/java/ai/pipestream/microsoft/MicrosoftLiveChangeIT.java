package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.grpc.ManagedChannel;
import io.grpc.MethodDescriptor;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.SecureRandom;
import java.time.Clock;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Runs {@link ChangeScenarios} against the real tenant with a notification
 * receiver reachable at a public URL, records every Graph exchange and every
 * notification Graph posted (scrubbed), and writes {@code build/live-changes.md}.
 *
 * <p>Runs only when the application credentials and site are set,
 * {@code MICROSOFT_LIVE_CHANGES=1}, and {@code MICROSOFT_NOTIFICATION_PUBLIC_URL}
 * is the public HTTPS base that forwards to {@code MICROSOFT_NOTIFICATION_PORT}
 * (default 8089) on this machine. It writes one file at the drive root and
 * three subscriptions, and removes them all.</p>
 */
class MicrosoftLiveChangeIT {

    /** Environment switch that allows the run. */
    static final String ENV_LIVE_CHANGES = "MICROSOFT_LIVE_CHANGES";
    /** Environment variable for the public HTTPS base of the receiver. */
    static final String ENV_PUBLIC_URL = "MICROSOFT_NOTIFICATION_PUBLIC_URL";
    /** Default receiver port. */
    static final int DEFAULT_PORT = 8089;
    /** Minutes the run waits for Graph to ask for a reauthorization; unset = two. */
    static final String ENV_LIFECYCLE_WAIT_MINUTES = "MICROSOFT_LIVE_LIFECYCLE_WAIT_MINUTES";

    private static final String TAG = "[MicrosoftLiveChangeIT] ";

    private record Row(String rpc, String label, String outcome, String detail) {
    }

    private final List<Row> rows = new ArrayList<>();
    private final Set<String> secrets = ConcurrentHashMap.newKeySet();
    private final List<JsonNode> pending = new ArrayList<>();
    private final List<JsonNode> lifecyclePending = new ArrayList<>();

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void everyChangeRpcAgainstTheLiveTenant() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID)),
                "no live Microsoft application credentials and site in the environment; skipping");
        assumeTrue("1".equals(System.getenv(ENV_LIVE_CHANGES)), ENV_LIVE_CHANGES + " is not 1");
        assumeTrue(notBlank(System.getenv(ENV_PUBLIC_URL)), ENV_PUBLIC_URL + " is not set");
        Path out = Path.of(System.getProperty("liveChanges.report", "build/live-changes.md"));
        List<String> failures = run(MicrosoftConnectorConfig.fromEnvironment(), out,
                ChangeRecordings.dir());
        assertThat(failures).as("change rpcs that failed against the tenant (see %s)", out).isEmpty();
    }

    List<String> run(MicrosoftConnectorConfig config, Path out, Path recordings) throws Exception {
        secrets.add(config.clientSecret());
        GraphAuth auth = new GraphAuth(config.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(auth.clientCredentials());
        secrets.add(token.get().accessToken());
        java.util.function.Supplier<String> bearer = () -> {
            GraphAuth.Token current = token.get();
            if (current.expired()) {
                try {
                    current = auth.clientCredentials();
                    secrets.add(current.accessToken());
                    token.set(current);
                } catch (IOException e) {
                    throw new IllegalStateException("token refresh failed", e);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    throw new IllegalStateException("token refresh interrupted", e);
                }
            }
            return current.accessToken();
        };
        String graphBaseUrl = GraphLinks.stripTrailingSlashes(config.graphBaseUrl());
        GraphClient direct = new GraphClient(graphBaseUrl, bearer, AuthMode.AUTH_MODE_APPLICATION);
        String site = "/sites/" + GraphPaths.site(config.siteId());
        String driveId = config.driveIds().isEmpty()
                ? direct.get(site + "/drive?$select=id").path("id").asText("")
                : config.driveIds().get(0);
        if (driveId.isBlank()) {
            throw new IllegalStateException("the site's default document library was not found");
        }

        String publicUrl = GraphLinks.stripTrailingSlashes(System.getenv(ENV_PUBLIC_URL));
        secrets.add(publicUrl);
        String suffix = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneOffset.UTC)
                .format(Instant.now()) + "-" + HexFormat.of().toHexDigits(new SecureRandom().nextInt());
        Instant startedAt = Instant.now().truncatedTo(java.time.temporal.ChronoUnit.SECONDS);
        ChangeScenarios.Backend backend = new ChangeScenarios.Backend(driveId, suffix, startedAt,
                publicUrl + NotificationReceiver.NOTIFICATIONS_PATH,
                publicUrl + NotificationReceiver.LIFECYCLE_PATH, config.siteId());

        Map<String, byte[]> blobs = new ConcurrentHashMap<>();
        RecordingScrubber scrubber = new RecordingScrubber(graphBaseUrl);
        RecordingTransport recorder = RecordingTransport.overHttp(scrubber, blobs, graphBaseUrl);
        GraphClient client = new GraphClient(graphBaseUrl, bearer, AuthMode.AUTH_MODE_APPLICATION,
                recorder);
        blobs.put(Recordings.sha256(ChangeRecordings.content(suffix)), ChangeRecordings.content(suffix));

        NotificationBus receiverBus = new NotificationBus();
        NotificationBus serviceBus = new NotificationBus();
        AtomicLong lastArrival = new AtomicLong(System.nanoTime());
        int port = Integer.parseInt(java.util.Objects.toString(
                System.getenv(NotificationReceiver.ENV_PORT), String.valueOf(DEFAULT_PORT)).trim());
        MicrosoftGrpcService microsoft = new MicrosoftGrpcService(config, new GraphFiles(client),
                MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES);
        MicrosoftChangeGrpcService changes = new MicrosoftChangeGrpcService(microsoft, serviceBus,
                true, backend.clientState(), Clock.systemUTC());
        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor).addService(changes)
                .addService(microsoft).build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        Instant started = Instant.now();
        try (NotificationReceiver receiver = NotificationReceiver.start(port, backend.clientState(),
                receiverBus, Clock.systemUTC())) {
            receiver.tap(envelope -> {
                synchronized (pending) {
                    if (lifecycle(envelope)) {
                        lifecyclePending.add(envelope);
                    } else {
                        pending.add(envelope);
                        lastArrival.set(System.nanoTime());
                    }
                }
            });
            SubscriptionRenewer renewer = ChangeScenarios.renewer(changes::subscriptions,
                    serviceBus, backend);
            changes.renewer(renewer);
            renewer.answerLifecycles();
            ChangeScenarios.Notifications live = new ChangeScenarios.Notifications() {
                @Override
                public void arrive(String stepId) {
                    // Graph posts on its own; the scenario waits for it.
                }

                @Override
                public void settle(Duration quiet) throws InterruptedException {
                    while (System.nanoTime() - lastArrival.get() < quiet.toNanos()) {
                        Thread.sleep(1000);
                    }
                    synchronized (pending) {
                        pending.clear();
                    }
                }
            };
            ChangeScenarios scenarios = new ChangeScenarios(
                    MicrosoftChangeServiceGrpc.newBlockingStub(channel),
                    MicrosoftChangeServiceGrpc.newStub(channel),
                    ai.pipestream.microsoft.v1.MicrosoftServiceGrpc.newBlockingStub(channel), client, backend, receiverBus,
                    serviceBus, live, renewer, () -> ChangeScenarios.renewer(
                            changes::subscriptions, new NotificationBus(), backend));
            String wait = System.getenv(ENV_LIFECYCLE_WAIT_MINUTES);
            if (notBlank(wait)) {
                scenarios.lifecycleWait(Duration.ofMinutes(Long.parseLong(wait.trim())));
            }
            clear(recordings);
            ObjectNode run = Recordings.JSON.createObjectNode()
                    .put("suffix", suffix)
                    .put("graphBaseUrl", graphBaseUrl)
                    .put("driveId", scrubber.text(driveId))
                    .put("startedAt", startedAt.toString())
                    .put("notificationUrl", scrubber.text(backend.notificationUrl()))
                    .put("lifecycleUrl", scrubber.text(backend.lifecycleUrl()))
                    .put("siteId", scrubber.text(backend.siteId()));
            NotificationBus.Listener listener = scenarios.listen();
            try {
                for (ChangeScenarios.Scenario scenario : scenarios.all()) {
                    recorder.drop();
                    WriteScenarios.Outcome outcome = scenarios.run(scenario);
                    String status = outcome.status().name();
                    String detail = outcome.detail();
                    if (outcome.status() == WriteScenarios.Status.PASS) {
                        try {
                            Path dir = recordings.resolve(scenario.id());
                            recorder.save(dir);
                            saveNotifications(scenario.id(), dir, scrubber);
                        } catch (IOException | IllegalStateException e) {
                            clear(recordings.resolve(scenario.id()));
                            status = "FAIL";
                            detail = "passed, but its recording was refused: " + e.getMessage();
                        }
                    }
                    rows.add(new Row(scenario.rpc(), scenario.label(), status, detail));
                    System.out.println(TAG + scrub(scenario.label() + ": " + status + " " + detail));
                }
            } finally {
                listener.close();
                scenarios.cleanUp();
                renewer.close();
            }
            Files.createDirectories(recordings);
            Files.writeString(recordings.resolve(Recordings.RUN_FILE),
                    Recordings.JSON.writerWithDefaultPrettyPrinter().writeValueAsString(run) + "\n",
                    StandardCharsets.UTF_8);
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            server.awaitTermination(10, TimeUnit.SECONDS);
            executor.shutdownNow();
        }
        for (MethodDescriptor<?, ?> method
                : MicrosoftChangeServiceGrpc.getServiceDescriptor().getMethods()) {
            String rpc = method.getBareMethodName();
            if (rpc.equals("WatchChanges") || rows.stream().anyMatch(row -> row.rpc().equals(rpc))) {
                continue;
            }
            rows.add(new Row(rpc, rpc, "FAIL", "not exercised by this test"));
        }
        String report = report(started);
        Files.createDirectories(out.toAbsolutePath().getParent());
        Files.writeString(out, report, StandardCharsets.UTF_8);
        System.out.println(report);
        List<String> failures = new ArrayList<>();
        rows.stream().filter(row -> row.outcome().equals("FAIL"))
                .forEach(row -> failures.add(scrub(row.label() + ": " + row.detail())));
        return failures;
    }

    /** Saves the notification bodies Graph posted since the last save, for steps that wait on them. */
    private void saveNotifications(String stepId, Path dir, RecordingScrubber scrubber)
            throws IOException {
        if (!ChangeRecordings.NOTIFIED_STEPS.contains(stepId)) {
            return;
        }
        ArrayNode bodies = Recordings.JSON.createArrayNode();
        synchronized (pending) {
            List<JsonNode> arrived = stepId.equals(ChangeScenarios.LIFECYCLE_STEP)
                    ? lifecyclePending : pending;
            arrived.forEach(body -> bodies.add(scrubber.json(body)));
            arrived.clear();
        }
        String text = Recordings.JSON.writerWithDefaultPrettyPrinter().writeValueAsString(bodies) + "\n";
        List<String> violations = RecordingScrubber.violations(text);
        if (!violations.isEmpty()) {
            throw new IllegalStateException("notifications of " + stepId
                    + " still hold tenant data: " + violations);
        }
        Files.createDirectories(dir);
        Files.writeString(dir.resolve(ChangeRecordings.NOTIFICATIONS_FILE), text,
                StandardCharsets.UTF_8);
    }

    /** Whether a posted body is a lifecycle notification. */
    private static boolean lifecycle(JsonNode envelope) {
        for (JsonNode item : envelope.path("value")) {
            if (!item.path("lifecycleEvent").asText("").isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private static void clear(Path directory) throws IOException {
        if (!Files.exists(directory)) {
            return;
        }
        try (Stream<Path> tree = Files.walk(directory)) {
            for (Path path : tree.sorted(Comparator.reverseOrder()).toList()) {
                Files.delete(path);
            }
        }
    }

    private String report(Instant started) {
        StringBuilder md = new StringBuilder("# Microsoft Graph live change run\n\n");
        md.append("- Run: ").append(started).append(" (")
                .append(Duration.between(started, Instant.now()).toSeconds()).append(" s)\n\n");
        md.append("| step | outcome | detail |\n|---|---|---|\n");
        for (Row row : rows) {
            md.append("| ").append(row.label().replace("|", "\\|")).append(" | ")
                    .append(row.outcome().equals("FAIL") ? "**FAIL**" : row.outcome()).append(" | ")
                    .append(String.valueOf(row.detail()).replace("|", "\\|").replace('\n', ' '))
                    .append(" |\n");
        }
        return scrub(md.toString());
    }

    private String scrub(String text) {
        String out = text;
        for (String secret : secrets) {
            if (notBlank(secret)) {
                out = out.replace(secret, "***");
            }
        }
        return out;
    }
}
