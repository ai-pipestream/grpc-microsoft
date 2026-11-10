package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.grpc.ManagedChannel;
import io.grpc.MethodDescriptor;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.SecureRandom;
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
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Supplier;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Runs {@link WriteScenarios} against the real tenant, records every Graph
 * exchange (scrubbed) under the recordings directory, and writes
 * {@code build/live-write.md}.
 *
 * <p>Runs only when {@code MICROSOFT_TENANT_ID}, {@code MICROSOFT_CLIENT_ID},
 * {@code MICROSOFT_CLIENT_SECRET} and {@code MICROSOFT_SITE_ID} are set and
 * {@code MICROSOFT_LIVE_WRITE=1}. Everything it writes sits in one new folder
 * at the drive root and one new list, both removed at the end. Run with
 * {@code ./gradlew :grpc-microsoft-service:liveWriteTest}.</p>
 */
class MicrosoftLiveWriteIT {

    /** Environment switch that allows writing to the tenant. */
    static final String ENV_LIVE_WRITE = "MICROSOFT_LIVE_WRITE";

    private static final String TAG = "[MicrosoftLiveWriteIT] ";

    /** The rpcs this test must exercise. */
    private static final Set<String> WRITE_RPCS = Set.of("UploadItem", "CreateFolder",
            "UpdateDriveItem", "DeleteDriveItem", "RestoreDriveItemVersion", "CopyDriveItem",
            "CreateListItem", "UpdateListItemFields", "DeleteListItem", "EnsureColumns",
            "EnsureList", "Mutate");

    private record Row(String rpc, String label, String outcome, String detail) {
    }

    private final List<Row> rows = new ArrayList<>();
    private final Set<String> secrets = ConcurrentHashMap.newKeySet();
    private String folderName;
    private String listName;
    private String recorded = "";

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void everyWriteRpcAgainstTheLiveTenant() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID)),
                "no live Microsoft application credentials and site in the environment; skipping");
        assumeTrue("1".equals(System.getenv(ENV_LIVE_WRITE)),
                ENV_LIVE_WRITE + " is not 1; not writing to the tenant");
        Path out = Path.of(System.getProperty("liveWrite.report", "build/live-write.md"));
        List<String> failures = run(MicrosoftConnectorConfig.fromEnvironment(), out,
                Recordings.dir());
        assertThat(failures).as("write rpcs that failed against the tenant (see %s)", out)
                .isEmpty();
    }

    /**
     * Runs every scenario, saves the recordings of those that passed, writes and
     * prints the report; answers the failed rows.
     */
    List<String> run(MicrosoftConnectorConfig config, Path out, Path recordings) throws Exception {
        secrets.add(config.clientSecret());
        GraphAuth auth = new GraphAuth(config.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(issue(auth));
        Supplier<String> bearer = () -> {
            GraphAuth.Token current = token.get();
            if (current.expired()) {
                try {
                    current = issue(auth);
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

        // Not recorded: these answers name the site and its libraries.
        GraphClient direct = new GraphClient(graphBaseUrl, bearer, AuthMode.AUTH_MODE_APPLICATION);
        String site = "/sites/" + GraphPaths.site(config.siteId());
        String siteId = direct.get(site + "?$select=id").path("id").asText("");
        String driveId = config.driveIds().isEmpty()
                ? direct.get(site + "/drive?$select=id").path("id").asText("")
                : config.driveIds().get(0);
        if (siteId.isBlank() || driveId.isBlank()) {
            throw new IllegalStateException("the site or its default document library was not found");
        }

        Map<String, byte[]> blobs = new ConcurrentHashMap<>();
        RecordingScrubber scrubber = new RecordingScrubber(graphBaseUrl);
        RecordingTransport recorder = RecordingTransport.overHttp(scrubber, blobs, graphBaseUrl);
        GraphClient client = new GraphClient(graphBaseUrl, bearer, AuthMode.AUTH_MODE_APPLICATION,
                recorder);
        String suffix = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneOffset.UTC)
                .format(Instant.now()) + "-" + HexFormat.of().toHexDigits(new SecureRandom().nextInt());
        MicrosoftGrpcService service = new MicrosoftGrpcService(config, new GraphFiles(client),
                MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES);
        service.mutateLingerMillis(WriteScenarios.MUTATE_LINGER_MILLIS);

        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor).addService(service)
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        Instant started = Instant.now();
        try {
            WriteScenarios scenarios = new WriteScenarios(MicrosoftServiceGrpc.newBlockingStub(channel),
                    MicrosoftServiceGrpc.newStub(channel), client,
                    new WriteScenarios.Backend(siteId, driveId, suffix, null, blobs));
            folderName = scenarios.folderName();
            listName = scenarios.listName();
            clear(recordings);
            ObjectNode run = Recordings.JSON.createObjectNode()
                    .put("suffix", suffix)
                    .put("graphBaseUrl", graphBaseUrl)
                    .put("siteId", scrubber.text(siteId))
                    .put("driveId", scrubber.text(driveId));
            int saved = 0;
            for (WriteScenarios.Scenario scenario : scenarios.all()) {
                recorder.drop();
                WriteScenarios.Outcome outcome = scenarios.run(scenario);
                String status = outcome.status().name();
                String detail = outcome.detail();
                if (outcome.status() == WriteScenarios.Status.PASS) {
                    try {
                        saved += recorder.save(recordings.resolve(scenario.id()));
                    } catch (IOException | IllegalStateException e) {
                        clear(recordings.resolve(scenario.id()));
                        status = "FAIL";
                        detail = "passed, but its recording was refused: " + e.getMessage();
                    }
                }
                rows.add(new Row(scenario.rpc(), scenario.label(), status, detail));
                System.out.println(TAG + scrub(scenario.label() + ": " + status + " " + detail));
            }
            Files.createDirectories(recordings);
            Files.writeString(recordings.resolve(Recordings.RUN_FILE),
                    Recordings.JSON.writerWithDefaultPrettyPrinter().writeValueAsString(run) + "\n",
                    StandardCharsets.UTF_8);
            recorded = saved + " exchanges under " + recordings;
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            server.awaitTermination(10, TimeUnit.SECONDS);
            executor.shutdownNow();
        }

        for (MethodDescriptor<?, ?> method : MicrosoftServiceGrpc.getServiceDescriptor().getMethods()) {
            String rpc = method.getBareMethodName();
            if (WRITE_RPCS.contains(rpc) && rows.stream().noneMatch(row -> row.rpc().equals(rpc))) {
                rows.add(new Row(rpc, rpc, "FAIL", "not exercised by this test"));
            }
        }
        String report = report(URI.create(graphBaseUrl).getHost(), started);
        Files.createDirectories(out.toAbsolutePath().getParent());
        Files.writeString(out, report, StandardCharsets.UTF_8);
        System.out.println(report);
        System.out.println(TAG + "report written to " + out.toAbsolutePath());

        List<String> failures = new ArrayList<>();
        rows.stream().filter(row -> row.outcome().equals("FAIL"))
                .forEach(row -> failures.add(scrub(row.label() + ": " + row.detail())));
        return failures;
    }

    private GraphAuth.Token issue(GraphAuth auth) throws IOException, InterruptedException {
        GraphAuth.Token issued = auth.clientCredentials();
        if (notBlank(issued.accessToken())) {
            secrets.add(issued.accessToken());
        }
        if (notBlank(issued.refreshToken())) {
            secrets.add(issued.refreshToken());
        }
        return issued;
    }

    /** Deletes a directory tree. */
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

    // ------------------------------------------------------------ report

    private String report(String host, Instant started) {
        StringBuilder md = new StringBuilder();
        md.append("# Microsoft Graph live write run\n\n");
        md.append("- Graph: ").append(host).append(" (application token, client credentials)\n");
        md.append("- Run: ").append(started).append(" (")
                .append(Duration.between(started, Instant.now()).toSeconds()).append(" s)\n");
        md.append("- Test folder: /").append(folderName).append("\n- Test list: ")
                .append(listName).append("\n");
        long failed = rows.stream().filter(row -> row.outcome().equals("FAIL")).count();
        md.append("- Rows: ").append(rows.size()).append(", failed ").append(failed).append("\n\n");
        md.append("## RPCs\n\n| rpc | outcome | detail |\n|---|---|---|\n");
        for (Row row : rows) {
            md.append("| ").append(cell(row.label())).append(" | ")
                    .append(row.outcome().equals("FAIL") ? "**FAIL**" : row.outcome())
                    .append(" | ").append(cell(row.detail())).append(" |\n");
        }
        md.append("\n## Recordings\n\n").append(recorded).append("\n");
        return scrub(md.toString());
    }

    private static String cell(String text) {
        String line = String.valueOf(text).replace('\n', ' ').replace('\r', ' ');
        return (line.length() > 400 ? line.substring(0, 400) + "..." : line).replace("|", "\\|");
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
