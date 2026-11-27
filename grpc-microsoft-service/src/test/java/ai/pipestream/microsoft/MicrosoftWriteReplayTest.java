package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.Message;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.TestFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Runs {@link WriteScenarios} against the exchanges {@link MicrosoftLiveWriteIT}
 * recorded from the real tenant. No network: any request that is not the next
 * recorded one fails its scenario. A scenario with no recording is skipped.
 */
class MicrosoftWriteReplayTest {

    /** System property naming a directory; when set, the records the scenarios read are written there. */
    static final String RECORDS_DIR = "replay.records.dir";

    private ExecutorService executor;
    private Server server;
    private ManagedChannel channel;

    @AfterEach
    void stop() {
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
    List<DynamicTest> scenarios() throws Exception {
        Path recordings = Recordings.dir();
        Path runFile = recordings.resolve(Recordings.RUN_FILE);
        boolean recorded = Files.isRegularFile(runFile);
        JsonNode run = recorded ? Recordings.JSON.readTree(runFile.toFile())
                : Recordings.JSON.createObjectNode();
        Map<String, byte[]> blobs = new ConcurrentHashMap<>();
        ReplayTransport transport = new ReplayTransport(blobs);
        List<String> termIds = new ArrayList<>();
        run.path("termIds").forEach(id -> termIds.add(id.asText()));
        WriteScenarios.Backend backend = new WriteScenarios.Backend(run.path("siteId").asText(),
                run.path("driveId").asText(), run.path("suffix").asText(), Duration.ofMillis(1),
                blobs, run.path("termSetId").asText(), termIds);
        WriteScenarios scenarios = recorded ? start(run, transport, backend)
                : new WriteScenarios(null, null, null, backend);

        List<DynamicTest> tests = new ArrayList<>();
        for (WriteScenarios.Scenario scenario : scenarios.all()) {
            tests.add(DynamicTest.dynamicTest(scenario.id(), () -> {
                Path directory = recordings.resolve(scenario.id());
                assumeTrue(recorded && Files.isDirectory(directory), "no recording of "
                        + scenario.id() + " yet under " + recordings
                        + "; record one with liveWriteTest");
                transport.begin(directory);
                WriteScenarios.Outcome outcome = scenarios.run(scenario);
                List<String> problems = transport.finish();
                assumeTrue(outcome.status() != WriteScenarios.Status.SKIPPED, outcome.detail());
                assertThat(problems).as("requests that differ from the recording").isEmpty();
                assertThat(outcome.status()).as(outcome.detail())
                        .isEqualTo(WriteScenarios.Status.PASS);
            }));
        }
        String export = System.getProperty(RECORDS_DIR, "");
        if (!export.isBlank()) {
            tests.add(DynamicTest.dynamicTest("write the records", () -> write(scenarios.records(),
                    Path.of(export))));
        }
        return tests;
    }

    /** Writes each kind of record to {@code {kind}.binpb}, length-delimited. */
    private static void write(Map<String, List<Message>> records, Path directory) throws IOException {
        assertThat(records).as("records the replayed scenarios read").isNotEmpty();
        Files.createDirectories(directory);
        for (Map.Entry<String, List<Message>> kind : records.entrySet()) {
            try (OutputStream out = Files.newOutputStream(directory.resolve(kind.getKey() + ".binpb"))) {
                for (Message message : kind.getValue()) {
                    message.writeDelimitedTo(out);
                }
            }
        }
    }

    private WriteScenarios start(JsonNode run, ReplayTransport transport,
            WriteScenarios.Backend backend) throws Exception {
        String graphBaseUrl = run.path("graphBaseUrl").asText();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("replay-tenant").clientId("replay-client").clientSecret("replay-secret")
                .siteId(backend.siteId()).graphBaseUrl(graphBaseUrl).authority(graphBaseUrl)
                .build();
        GraphClient client = new GraphClient(graphBaseUrl, () -> "replay-token",
                AuthMode.AUTH_MODE_APPLICATION, transport);
        MicrosoftGrpcService service = new MicrosoftGrpcService(config, new GraphFiles(client),
                MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES);
        service.mutateLingerMillis(WriteScenarios.MUTATE_LINGER_MILLIS);
        service.sharePointTokens((session, resource) -> () -> "replay-token");
        executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor).addService(service).build()
                .start();
        channel = InProcessChannelBuilder.forName(name).build();
        return new WriteScenarios(MicrosoftServiceGrpc.newBlockingStub(channel),
                MicrosoftServiceGrpc.newStub(channel), client, backend);
    }
}
