package ai.pipestream.microsoft;

import ai.pipestream.microsoft.admin.MicrosoftAdminGrpcService;
import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.ListSiteAppGrantsRequest;
import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.RevokeSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.SiteAppGrant;
import ai.pipestream.microsoft.v1.SiteAppGrantState;
import ai.pipestream.microsoft.v1.SiteAppRole;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;
import java.util.stream.Stream;

/**
 * The admin service scenarios that run against the real tenant, where their
 * Graph exchanges are recorded, and against those recordings.
 */
final class AdminGrantScenarios {

    /** Where the admin recordings live, relative to the module directory. */
    static final String DEFAULT_DIR = "src/test/resources/recordings/admin";
    /** System property that overrides {@link #DEFAULT_DIR}. */
    static final String DIR_PROPERTY = "adminRecordings.dir";
    /** An application id no application has; the scrubber leaves it as it is. */
    static final String ABSENT_APP = RecordingScrubber.GUID_PREFIX + "999999999999";

    /** How a scenario ended. */
    enum Status { PASS, FAIL, SKIPPED }

    /**
     * What the scenarios work on.
     *
     * @param siteId the site whose grants are read and changed
     * @param granteeId the application that already holds one grant on the site
     * @param selfId the application the lifecycle scenario grants and revokes
     * @param selfName the display name the lifecycle scenario gives that grant
     * @param change whether the lifecycle scenario may change grants
     */
    record Backend(String siteId, String granteeId, String selfId, String selfName, boolean change) {
    }

    /** One scenario: the directory of its recording, the rpc it proves and what it does. */
    record Scenario(String id, String rpc, String label, Body body) {
    }

    /** The end of one scenario. */
    record Outcome(Status status, String detail) {
    }

    /** One scenario's outcome in a run. */
    record Row(Scenario scenario, Outcome outcome) {
    }

    /** What a scenario does; answers a detail without tenant data. */
    @FunctionalInterface
    interface Body {
        String run() throws Exception;
    }

    /** Ends a scenario as skipped. */
    private static final class Skip extends RuntimeException {
        private static final long serialVersionUID = 1L;

        Skip(String message) {
            super(message);
        }
    }

    private final MicrosoftAdminServiceGrpc.MicrosoftAdminServiceBlockingStub stub;
    private final Backend backend;

    AdminGrantScenarios(MicrosoftAdminServiceGrpc.MicrosoftAdminServiceBlockingStub stub, Backend backend) {
        this.stub = stub;
        this.backend = backend;
    }

    /** The admin recordings directory of this build. */
    static Path dir() {
        return Path.of(System.getProperty(DIR_PROPERTY, DEFAULT_DIR));
    }

    /** Every scenario, in the order they run. */
    List<Scenario> all() {
        return List.of(
                new Scenario("list-site-app-grants", "ListSiteAppGrants", "ListSiteAppGrants",
                        this::listGrants),
                new Scenario("ensure-site-app-grant-dry-run", "EnsureSiteAppGrant",
                        "EnsureSiteAppGrant (dry run)", this::dryRuns),
                new Scenario("site-app-grant-lifecycle", "RevokeSiteAppGrant",
                        "EnsureSiteAppGrant and RevokeSiteAppGrant (create, same, change, revoke)",
                        this::lifecycle));
    }

    /** Runs one scenario. */
    Outcome run(Scenario scenario) {
        try {
            return new Outcome(Status.PASS, scenario.body().run());
        } catch (Skip e) {
            return new Outcome(Status.SKIPPED, e.getMessage());
        } catch (Exception | AssertionError e) {
            return new Outcome(Status.FAIL, e.getClass().getSimpleName() + ": " + e.getMessage());
        }
    }

    private static void check(boolean ok, String message) {
        if (!ok) {
            throw new AssertionError(message);
        }
    }

    private MicrosoftAdminServiceGrpc.MicrosoftAdminServiceBlockingStub call() {
        return stub.withDeadlineAfter(120, TimeUnit.SECONDS);
    }

    private List<SiteAppGrant> list() {
        List<SiteAppGrant> grants = new ArrayList<>();
        call().listSiteAppGrants(ListSiteAppGrantsRequest.newBuilder().setSiteId(backend.siteId()).build())
                .forEachRemaining(response -> grants.add(response.getGrant()));
        return grants;
    }

    private List<SiteAppGrant> of(List<SiteAppGrant> grants, String appId) {
        return grants.stream().filter(grant -> grant.getAppId().equalsIgnoreCase(appId)).toList();
    }

    /** The one grant the test application holds, with exactly one known role. */
    private SiteAppGrant held(List<SiteAppGrant> grants) {
        List<SiteAppGrant> own = of(grants, backend.granteeId());
        check(own.size() == 1, "the test application holds " + own.size() + " grants on the site, not 1");
        check(own.get(0).getRolesCount() == 1 && own.get(0).getOtherRolesCount() == 0,
                "the grant of the test application does not have exactly one known role");
        return own.get(0);
    }

    private String listGrants() {
        List<SiteAppGrant> grants = list();
        check(!grants.isEmpty(), "the site has no application grants");
        for (SiteAppGrant grant : grants) {
            check(!grant.getId().isBlank() && !grant.getAppId().isBlank(), "a grant has no id or app id");
            check(grant.getSiteId().equals(backend.siteId()), "a grant names another site");
            check(grant.getRolesCount() + grant.getOtherRolesCount() > 0, "a grant has no role");
        }
        return grants.size() + " grant(s); the test app holds " + held(grants).getRolesList();
    }

    private EnsureSiteAppGrantRequest.Builder ensure(String appId, SiteAppRole role) {
        return EnsureSiteAppGrantRequest.newBuilder().setSiteId(backend.siteId()).setAppId(appId)
                .setRole(role);
    }

    private String dryRuns() {
        List<SiteAppGrant> before = list();
        SiteAppGrant own = held(before);
        SiteAppRole role = own.getRoles(0);

        EnsureSiteAppGrantResponse same = call().ensureSiteAppGrant(
                ensure(backend.granteeId(), role).setDryRun(true).build());
        check(same.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS,
                "the held role answered " + same.getState());
        check(same.getGrant().equals(own), "the held role did not answer the grant");

        SiteAppRole other = role == SiteAppRole.SITE_APP_ROLE_READ
                ? SiteAppRole.SITE_APP_ROLE_WRITE : SiteAppRole.SITE_APP_ROLE_READ;
        EnsureSiteAppGrantResponse change = call().ensureSiteAppGrant(
                ensure(backend.granteeId(), other).setDryRun(true).build());
        check(change.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_UPDATE,
                "another role answered " + change.getState());
        check(change.getPreviousRolesList().equals(List.of(role)), "previous_roles is not the held role");

        EnsureSiteAppGrantResponse absent = call().ensureSiteAppGrant(
                ensure(ABSENT_APP, SiteAppRole.SITE_APP_ROLE_READ).setDryRun(true).build());
        check(absent.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_CREATE,
                "an application without a grant answered " + absent.getState());
        check(!absent.hasGrant(), "a dry run answered a grant that does not exist");

        check(list().equals(before), "the grants changed during the dry runs");
        return "held " + role + " -> EXISTS, " + other + " -> WOULD_UPDATE, unknown app -> WOULD_CREATE;"
                + " grants unchanged";
    }

    private String lifecycle() {
        if (!backend.change()) {
            throw new Skip("changes grants; runs only when the run allows it");
        }
        String self = backend.selfId();
        List<SiteAppGrant> before = list();
        check(of(before, self).isEmpty(), "the application already holds a grant before the run");
        EnsureSiteAppGrantRequest.Builder ensure = ensure(self, SiteAppRole.SITE_APP_ROLE_READ)
                .setAppDisplayName(backend.selfName());
        int revoked;
        try {
            EnsureSiteAppGrantResponse created = call().ensureSiteAppGrant(ensure.build());
            check(created.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_CREATED,
                    "a new grant answered " + created.getState());
            check(created.getGrant().getRolesList().equals(List.of(SiteAppRole.SITE_APP_ROLE_READ)),
                    "the new grant does not hold read");
            check(list().contains(created.getGrant()), "the new grant is not listed");

            EnsureSiteAppGrantResponse again = call().ensureSiteAppGrant(ensure.build());
            check(again.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS,
                    "the same role again answered " + again.getState());

            EnsureSiteAppGrantResponse updated = call().ensureSiteAppGrant(
                    ensure.setRole(SiteAppRole.SITE_APP_ROLE_WRITE).build());
            check(updated.getState() == SiteAppGrantState.SITE_APP_GRANT_STATE_UPDATED,
                    "another role answered " + updated.getState());
            check(updated.getPreviousRolesList().equals(List.of(SiteAppRole.SITE_APP_ROLE_READ)),
                    "previous_roles is not read");
            List<SiteAppGrant> own = of(list(), self);
            check(own.size() == 1 && own.get(0).getRolesList().equals(List.of(SiteAppRole.SITE_APP_ROLE_WRITE)),
                    "the grant does not hold write after the change");
        } finally {
            revoked = call().revokeSiteAppGrant(RevokeSiteAppGrantRequest.newBuilder()
                    .setSiteId(backend.siteId()).setAppId(self).build()).getRevoked();
        }
        check(revoked == 1, revoked + " grants were revoked, not 1");
        check(list().equals(before), "the grants are not back to what they were");
        return "CREATED, EXISTS, UPDATED, 1 revoked; grants are back to what they were";
    }

    // ------------------------------------------------------------ record and replay

    /** A running in-process admin service and the scenarios that call it. */
    private static final class Running implements AutoCloseable {
        private final ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        private final Server server;
        private final ManagedChannel channel;
        private final AdminGrantScenarios scenarios;

        Running(GraphClient graph, Backend backend) throws IOException {
            String name = InProcessServerBuilder.generateName();
            server = InProcessServerBuilder.forName(name).executor(executor)
                    .addService(new MicrosoftAdminGrpcService(graph)).build().start();
            channel = InProcessChannelBuilder.forName(name).build();
            scenarios = new AdminGrantScenarios(MicrosoftAdminServiceGrpc.newBlockingStub(channel), backend);
        }

        @Override
        public void close() {
            channel.shutdownNow();
            server.shutdownNow();
            executor.shutdownNow();
        }
    }

    /**
     * Runs every scenario through {@code real}, saves the scrubbed exchanges of
     * those that passed under {@code recordings}, replacing what was there, and
     * writes the run file.
     *
     * @return one row per scenario
     */
    static List<Row> record(String graphBaseUrl, Supplier<String> bearer, GraphTransport real,
            Backend backend, Path recordings) throws IOException {
        RecordingScrubber scrubber = new RecordingScrubber(graphBaseUrl);
        RecordingTransport recorder = new RecordingTransport(real, scrubber, new ConcurrentHashMap<>(),
                graphBaseUrl);
        List<Row> rows = new ArrayList<>();
        clear(recordings);
        try (Running running = new Running(new GraphClient(graphBaseUrl, bearer,
                AuthMode.AUTH_MODE_APPLICATION, recorder), backend)) {
            for (Scenario scenario : running.scenarios.all()) {
                recorder.drop();
                Outcome outcome = running.scenarios.run(scenario);
                if (outcome.status() == Status.PASS) {
                    try {
                        recorder.save(recordings.resolve(scenario.id()));
                    } catch (IOException | IllegalStateException e) {
                        clear(recordings.resolve(scenario.id()));
                        outcome = new Outcome(Status.FAIL, "passed, but its recording was refused: "
                                + e.getMessage());
                    }
                }
                rows.add(new Row(scenario, outcome));
            }
        }
        // The display name goes out inside an application identity, where the scrubber replaces it.
        ObjectNode named = Recordings.JSON.createObjectNode();
        named.putObject("application").put("displayName", backend.selfName());
        ObjectNode run = Recordings.JSON.createObjectNode()
                .put("graphBaseUrl", graphBaseUrl)
                .put("siteId", scrubber.text(backend.siteId()))
                .put("granteeId", scrubber.text(backend.granteeId()))
                .put("selfId", scrubber.text(backend.selfId()))
                .put("selfName", scrubber.json(named).path("application").path("displayName").asText())
                .put("change", backend.change());
        Files.createDirectories(recordings);
        Files.writeString(recordings.resolve(Recordings.RUN_FILE),
                Recordings.JSON.writerWithDefaultPrettyPrinter().writeValueAsString(run) + "\n",
                StandardCharsets.UTF_8);
        return rows;
    }

    /** One scenario of a recorded run and what replaying it found. */
    record Replayed(Scenario scenario, boolean recorded, Outcome outcome, List<String> problems) {
    }

    /**
     * Runs every scenario against the exchanges under {@code recordings}, with no
     * network; a scenario without a recording is not run.
     */
    static List<Replayed> replay(Path recordings) throws IOException {
        Path runFile = recordings.resolve(Recordings.RUN_FILE);
        if (!Files.isRegularFile(runFile)) {
            Backend none = new Backend("", "", "", "", false);
            return new AdminGrantScenarios(null, none).all().stream()
                    .map(scenario -> new Replayed(scenario, false, null, List.of())).toList();
        }
        JsonNode run = Recordings.JSON.readTree(runFile.toFile());
        Backend backend = new Backend(run.path("siteId").asText(), run.path("granteeId").asText(),
                run.path("selfId").asText(), run.path("selfName").asText(), run.path("change").asBoolean());
        ReplayTransport transport = new ReplayTransport(new ConcurrentHashMap<>());
        List<Replayed> replayed = new ArrayList<>();
        try (Running running = new Running(new GraphClient(run.path("graphBaseUrl").asText(),
                () -> "replay-token", AuthMode.AUTH_MODE_APPLICATION, transport), backend)) {
            for (Scenario scenario : running.scenarios.all()) {
                Path directory = recordings.resolve(scenario.id());
                if (!Files.isDirectory(directory)) {
                    replayed.add(new Replayed(scenario, false, null, List.of()));
                    continue;
                }
                transport.begin(directory);
                Outcome outcome = running.scenarios.run(scenario);
                replayed.add(new Replayed(scenario, true, outcome, transport.finish()));
            }
        }
        return replayed;
    }

    /** What under {@code recordings} still looks like tenant data, by file. */
    static List<String> violations(Path recordings) throws IOException {
        List<String> found = new ArrayList<>();
        if (!Files.isDirectory(recordings)) {
            return found;
        }
        try (Stream<Path> tree = Files.walk(recordings)) {
            for (Path file : tree.filter(Files::isRegularFile).sorted().toList()) {
                RecordingScrubber.violations(Files.readString(file, StandardCharsets.UTF_8))
                        .forEach(violation -> found.add(recordings.relativize(file) + ": " + violation));
            }
        }
        return found;
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
}
