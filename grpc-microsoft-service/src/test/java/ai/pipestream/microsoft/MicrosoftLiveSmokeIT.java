package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.GetSiteRequest;
import ai.pipestream.microsoft.v1.ListDrivesRequest;
import ai.pipestream.microsoft.v1.ListSitesRequest;
import ai.pipestream.microsoft.v1.ListSitesResponse;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.Site;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.Test;

import java.util.concurrent.Executors;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Cheap read-only probes against a real Microsoft 365 tenant. Gated on
 * Entra <em>application</em> credentials ({@code MICROSOFT_TENANT_ID},
 * {@code MICROSOFT_CLIENT_ID}, {@code MICROSOFT_CLIENT_SECRET}) — a user
 * mailbox login is not enough. Client-credentials tokens have no {@code /me},
 * so this class never calls {@code GetMe}; it reads the site named by
 * {@code MICROSOFT_SITE_ID} (or lists sites when none is set) and its drives. Excluded from the default {@code test}
 * task; run with {@code ./gradlew :grpc-microsoft-service:liveSmokeTest}.
 * Never prints the secret. Never runs a full {@code Sync} crawl.
 */
class MicrosoftLiveSmokeIT {

    private static final MicrosoftValidator VALIDATOR = MicrosoftValidator.create();

    @Test
    void clientCredentialsThenSiteAndDrives() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID)));
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID)));
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET)),
                "no live Microsoft application credentials; skipping");

        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment();
        GraphAuth auth = new GraphAuth(config.authConfig());
        GraphAuth.Token token = auth.clientCredentials();
        assertThat(token.accessToken()).isNotBlank();
        assertThat(token.expired()).isFalse();

        GraphFiles files = new GraphFiles(new GraphClient(config.graphBaseUrl(), token::accessToken));
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name)
                .executor(Executors.newVirtualThreadPerTaskExecutor())
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES))
                .build()
                .start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                    MicrosoftServiceGrpc.newBlockingStub(channel);
            String siteId = first(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID), "");
            if (siteId.isBlank()) {
                // Needs tenant-wide site access; an app with Sites.Selected sets the site id.
                ListSitesResponse sites = stub.listSites(ListSitesRequest.newBuilder()
                        .setLimit(5)
                        .build());
                assertThat(sites.getSitesList()).allSatisfy(site -> {
                    assertThat(site.getId()).isNotBlank();
                    assertThat(VALIDATOR.validate(site).violations()).isEmpty();
                });
                System.out.println("[MicrosoftLiveSmokeIT] ListSites -> " + sites.getSitesCount()
                        + " site(s)");
                if (sites.getSitesCount() == 0) {
                    return;
                }
                siteId = sites.getSites(0).getId();
            } else {
                Site site = stub.getSite(GetSiteRequest.newBuilder().setSiteId(siteId).build()).getSite();
                assertThat(site.getId()).isNotBlank();
                assertThat(VALIDATOR.validate(site).violations()).isEmpty();
                System.out.println("[MicrosoftLiveSmokeIT] GetSite -> the configured site");
            }
            var drives = stub.listDrives(ListDrivesRequest.newBuilder()
                    .setSiteId(siteId)
                    .build());
            assertThat(drives.getDrivesList()).isNotEmpty().allSatisfy(drive -> {
                assertThat(drive.getId()).isNotBlank();
                assertThat(VALIDATOR.validate(drive).violations()).isEmpty();
            });
            System.out.println("[MicrosoftLiveSmokeIT] ListDrives -> " + drives.getDrivesCount()
                    + " drive(s)");
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
        }
    }

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    private static String first(String preferred, String fallback) {
        if (preferred != null && !preferred.isBlank()) {
            return preferred.trim();
        }
        return fallback == null ? "" : fallback;
    }
}
