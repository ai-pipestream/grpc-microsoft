package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import ai.pipestream.microsoft.admin.MicrosoftAdminConfig;
import ai.pipestream.microsoft.admin.MicrosoftAdminServer;
import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.SiteAppGrantState;
import ai.pipestream.microsoft.v1.SiteAppRole;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.net.http.HttpClient;
import java.nio.file.Path;
import java.time.Duration;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;
import org.junit.jupiter.api.Test;

/**
 * Runs {@link AdminGrantScenarios} against the real tenant with the admin
 * identity and records every Graph exchange, scrubbed, under the admin
 * recordings directory. It reads the application grants of
 * {@code MICROSOFT_SITE_ID} and dry-runs an ensure for
 * {@code MICROSOFT_CLIENT_ID}; that part changes nothing in the tenant. With
 * {@code MICROSOFT_ADMIN_LIVE_CHANGE=1} it also grants, changes and revokes a
 * grant for the admin application itself. With
 * {@code MICROSOFT_ADMIN_LIVE_SET_ROLE} a second test sets the role of
 * {@code MICROSOFT_CLIENT_ID} on the site; that one is not recorded. Prints
 * counts and roles only. Run with
 * {@code ./gradlew :grpc-microsoft-admin:liveAdminTest}.
 */
class MicrosoftAdminLiveIT {

    private static final String TAG = "[MicrosoftAdminLiveIT] ";

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    private static boolean adminIdentity() {
        return notBlank(System.getenv(MicrosoftAdminConfig.ENV_TENANT_ID))
                && notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_ID))
                && (notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_CERT_PEM))
                || notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_SECRET)))
                && notBlank(System.getenv("MICROSOFT_SITE_ID"))
                && notBlank(System.getenv("MICROSOFT_CLIENT_ID"));
    }

    @Test
    void everyAdminRpcAgainstTheLiveTenant() throws Exception {
        assumeTrue(adminIdentity(), "no admin identity, site and grantee in the environment; skipping");
        MicrosoftAdminConfig config = MicrosoftAdminConfig.fromEnvironment();
        String graphBaseUrl = GraphLinks.stripTrailingSlashes(config.graphBaseUrl());
        Supplier<String> bearer = new GraphAuth(config.authConfig()).tokens(GraphAuth.GRAPH_SCOPE);

        // Not recorded: this answer names the site.
        String siteId = new GraphClient(graphBaseUrl, bearer, AuthMode.AUTH_MODE_APPLICATION)
                .get("/sites/" + GraphPaths.site(System.getenv("MICROSOFT_SITE_ID").trim()) + "?$select=id")
                .path("id").asText("");
        assertThat(siteId).as("the id of the test site").isNotBlank();

        AdminGrantScenarios.Backend backend = new AdminGrantScenarios.Backend(siteId,
                System.getenv("MICROSOFT_CLIENT_ID").trim(),
                System.getenv(MicrosoftAdminConfig.ENV_CLIENT_ID).trim(), "pipestream admin live test",
                "1".equals(System.getenv("MICROSOFT_ADMIN_LIVE_CHANGE")));
        Path recordings = AdminGrantScenarios.dir();
        List<AdminGrantScenarios.Row> rows = AdminGrantScenarios.record(graphBaseUrl, bearer,
                HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(20)).build()::send, backend,
                recordings);
        for (AdminGrantScenarios.Row row : rows) {
            System.out.println(TAG + row.scenario().label() + ": " + row.outcome().status() + " "
                    + row.outcome().detail());
        }
        assertThat(AdminGrantScenarios.violations(recordings)).as("tenant data in the recordings").isEmpty();
        System.out.println(TAG + "recordings written to " + recordings);
        assertThat(rows).as("scenarios that failed against the tenant")
                .noneMatch(row -> row.outcome().status() == AdminGrantScenarios.Status.FAIL);
    }

    @Test
    void theRoleOfTheTestApplicationIsSet() throws Exception {
        String wanted = System.getenv("MICROSOFT_ADMIN_LIVE_SET_ROLE");
        String granteeId = System.getenv("MICROSOFT_CLIENT_ID");
        assumeTrue(adminIdentity(), "no admin identity, site and grantee in the environment; skipping");
        assumeTrue(notBlank(wanted), "MICROSOFT_ADMIN_LIVE_SET_ROLE is not set; not changing grants");
        SiteAppRole role = SiteAppRole.valueOf("SITE_APP_ROLE_" + wanted.trim().toUpperCase(Locale.ROOT));
        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(MicrosoftAdminServer.service(MicrosoftAdminConfig.fromEnvironment()))
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        try {
            EnsureSiteAppGrantResponse response = MicrosoftAdminServiceGrpc.newBlockingStub(channel)
                    .withDeadlineAfter(120, TimeUnit.SECONDS)
                    .ensureSiteAppGrant(EnsureSiteAppGrantRequest.newBuilder()
                            .setSiteId(System.getenv("MICROSOFT_SITE_ID")).setAppId(granteeId)
                            .setRole(role).build());
            System.out.println(TAG + "EnsureSiteAppGrant [test app, " + role + "] -> " + response.getState()
                    + ", was " + response.getPreviousRolesList());
            assertThat(response.getState()).isIn(SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS,
                    SiteAppGrantState.SITE_APP_GRANT_STATE_UPDATED);
            assertThat(response.getGrant().getRolesList()).containsExactly(role);
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            executor.shutdownNow();
        }
    }
}
