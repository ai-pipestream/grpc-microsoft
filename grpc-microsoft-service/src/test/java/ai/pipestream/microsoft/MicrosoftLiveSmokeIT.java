package ai.pipestream.microsoft;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfEnvironmentVariable;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * One cheap Graph read. Skipped unless tenant/client/secret are set; the
 * default test task excludes this class.
 */
class MicrosoftLiveSmokeIT {

    @Test
    @EnabledIfEnvironmentVariable(named = "MICROSOFT_TENANT_ID", matches = ".+")
    void getMe() throws Exception {
        assumeTrue(notBlank(System.getenv("MICROSOFT_CLIENT_ID")));
        assumeTrue(notBlank(System.getenv("MICROSOFT_CLIENT_SECRET")));
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment();
        GraphAuth auth = new GraphAuth(config.authConfig());
        GraphAuth.Token token = auth.clientCredentials();
        GraphFiles files = new GraphFiles(new GraphClient(config.graphBaseUrl(), token::accessToken));
        assertThat(files.me().path("id").asText()).isNotBlank();
    }

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }
}
