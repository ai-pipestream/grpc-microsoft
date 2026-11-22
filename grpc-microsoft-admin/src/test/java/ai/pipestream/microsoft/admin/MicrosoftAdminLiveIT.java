package ai.pipestream.microsoft.admin;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import ai.pipestream.microsoft.v1.EnsureSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.ListSiteAppGrantsRequest;
import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.SiteAppGrant;
import ai.pipestream.microsoft.v1.SiteAppGrantState;
import ai.pipestream.microsoft.v1.SiteAppRole;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import org.junit.jupiter.api.Test;

/**
 * Reads the application grants of {@code MICROSOFT_SITE_ID} with the admin
 * identity and dry-runs an ensure for {@code MICROSOFT_CLIENT_ID}. Changes
 * nothing in the tenant. Prints counts and roles only. Run with
 * {@code ./gradlew :grpc-microsoft-admin:liveAdminTest}.
 */
class MicrosoftAdminLiveIT {

    private static final String TAG = "[MicrosoftAdminLiveIT] ";

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void grantsOfTheTestSiteAreReadAndAnEnsureIsDryRun() throws Exception {
        String siteId = System.getenv("MICROSOFT_SITE_ID");
        String granteeId = System.getenv("MICROSOFT_CLIENT_ID");
        assumeTrue(notBlank(System.getenv(MicrosoftAdminConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_ID))
                        && (notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_CERT_PEM))
                        || notBlank(System.getenv(MicrosoftAdminConfig.ENV_CLIENT_SECRET)))
                        && notBlank(siteId) && notBlank(granteeId),
                "no admin identity, site and grantee in the environment; skipping");

        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(MicrosoftAdminServer.service(MicrosoftAdminConfig.fromEnvironment()))
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        try {
            MicrosoftAdminServiceGrpc.MicrosoftAdminServiceBlockingStub stub =
                    MicrosoftAdminServiceGrpc.newBlockingStub(channel).withDeadlineAfter(120, TimeUnit.SECONDS);
            List<SiteAppGrant> grants = new ArrayList<>();
            stub.listSiteAppGrants(ListSiteAppGrantsRequest.newBuilder().setSiteId(siteId).build())
                    .forEachRemaining(response -> grants.add(response.getGrant()));
            assertThat(grants).isNotEmpty().allSatisfy(grant -> {
                assertThat(grant.getId()).isNotBlank();
                assertThat(grant.getAppId()).isNotBlank();
                assertThat(grant.getRolesCount() + grant.getOtherRolesCount()).isPositive();
            });
            List<SiteAppGrant> own = grants.stream()
                    .filter(grant -> grant.getAppId().equalsIgnoreCase(granteeId.trim())).toList();
            System.out.println(TAG + "ListSiteAppGrants -> " + grants.size() + " grant(s); the test app holds "
                    + own.stream().flatMap(grant -> grant.getRolesList().stream()).map(Enum::name).toList());
            assertThat(own).as("grants of the test app on the test site").hasSize(1);
            SiteAppRole held = own.get(0).getRoles(0);

            EnsureSiteAppGrantResponse same = stub.ensureSiteAppGrant(EnsureSiteAppGrantRequest.newBuilder()
                    .setSiteId(siteId).setAppId(granteeId).setRole(held).setDryRun(true).build());
            System.out.println(TAG + "EnsureSiteAppGrant [dry run, held role] -> " + same.getState());
            assertThat(same.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS);

            SiteAppRole other = held == SiteAppRole.SITE_APP_ROLE_READ
                    ? SiteAppRole.SITE_APP_ROLE_WRITE : SiteAppRole.SITE_APP_ROLE_READ;
            EnsureSiteAppGrantResponse change = stub.ensureSiteAppGrant(EnsureSiteAppGrantRequest.newBuilder()
                    .setSiteId(siteId).setAppId(granteeId).setRole(other).setDryRun(true).build());
            System.out.println(TAG + "EnsureSiteAppGrant [dry run, other role] -> " + change.getState());
            assertThat(change.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_UPDATE);
            assertThat(change.getPreviousRolesList()).containsExactly(held);

            EnsureSiteAppGrantResponse absent = stub.ensureSiteAppGrant(EnsureSiteAppGrantRequest.newBuilder()
                    .setSiteId(siteId).setAppId("00000000-0000-0000-0000-000000000001")
                    .setRole(SiteAppRole.SITE_APP_ROLE_READ).setDryRun(true).build());
            System.out.println(TAG + "EnsureSiteAppGrant [dry run, unknown app] -> " + absent.getState());
            assertThat(absent.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_CREATE);

            List<SiteAppGrant> after = new ArrayList<>();
            stub.listSiteAppGrants(ListSiteAppGrantsRequest.newBuilder().setSiteId(siteId).build())
                    .forEachRemaining(response -> after.add(response.getGrant()));
            assertThat(after).as("grants after the dry runs").isEqualTo(grants);
            System.out.println(TAG + "grants unchanged after the dry runs");
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            executor.shutdownNow();
        }
    }
}
