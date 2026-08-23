package ai.pipestream.microsoft.connector;

import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ConnectorCustomConfigTest {

    @Test
    void blankJsonUsesDefaults() {
        ConnectorCustomConfig config = ConnectorCustomConfig.parse("", null);

        assertThat(config.target()).isEqualTo(ConnectorCustomConfig.DEFAULT_TARGET);
        assertThat(config.plaintext()).isTrue();
        assertThat(config.driveIds()).isEmpty();
        assertThat(config.folderPath()).isEqualTo("/");
        assertThat(config.includeContent()).isFalse();
    }

    @Test
    void parsesAdminJson() {
        ConnectorCustomConfig config = ConnectorCustomConfig.parse("""
                {
                  "target": "graph-proxy:9096",
                  "plaintext": false,
                  "driveIds": ["d1", "d2"],
                  "folderPath": "/Shared",
                  "includeContent": true
                }
                """, "https://contoso.sharepoint.com");

        assertThat(config.target()).isEqualTo("graph-proxy:9096");
        assertThat(config.plaintext()).isFalse();
        assertThat(config.driveIds()).containsExactly("d1", "d2");
        assertThat(config.folderPath()).isEqualTo("/Shared");
        assertThat(config.includeContent()).isTrue();
    }

    @Test
    void rejectsNonJson() {
        assertThatThrownBy(() -> ConnectorCustomConfig.parse("not-json", null))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("not JSON");
    }

    @Test
    void hostPortDatasourceUrlIsTheGrpcTarget() {
        ConnectorCustomConfig config = ConnectorCustomConfig.parse("", "127.0.0.1:9096");
        assertThat(config.target()).isEqualTo("127.0.0.1:9096");
        assertThat(ConnectorCustomConfig.parse("", "grpc://graph-proxy:9096").target())
                .isEqualTo("graph-proxy:9096");
        assertThat(ConnectorCustomConfig.parse("", "https://contoso.sharepoint.com").target())
                .isEqualTo(ConnectorCustomConfig.DEFAULT_TARGET);
    }
}
