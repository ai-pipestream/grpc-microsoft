package ai.pipestream.microsoft;

import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class MicrosoftConnectorConfigTest {

    private static MicrosoftConnectorConfig.Builder valid() {
        return MicrosoftConnectorConfig.builder()
                .tenantId("tenant")
                .clientId("client")
                .clientSecret("secret");
    }

    @Test
    void defaultsAndRedaction() {
        MicrosoftConnectorConfig config = valid().build();
        assertThat(config.folderPath()).isEqualTo("/");
        assertThat(config.graphBaseUrl()).isEqualTo("https://graph.microsoft.com/v1.0");
        assertThat(config.hasDriveAllowlist()).isFalse();
        assertThat(config.toString()).doesNotContain("secret");
        assertThat(config.toString()).contains("***");
    }

    @Test
    void requiresTenantClientAndSecret() {
        assertThatThrownBy(() -> valid().tenantId("").build())
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("MICROSOFT_TENANT_ID");
        assertThatThrownBy(() -> valid().clientId(null).build())
                .isInstanceOf(IllegalArgumentException.class);
        assertThatThrownBy(() -> valid().clientSecret(" ").build())
                .isInstanceOf(IllegalArgumentException.class);
    }

    @Test
    void environment() {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment(Map.of(
                "MICROSOFT_TENANT_ID", "t",
                "MICROSOFT_CLIENT_ID", "c",
                "MICROSOFT_CLIENT_SECRET", "s",
                "MICROSOFT_DRIVE_IDS", "d1, d2",
                "MICROSOFT_FOLDER_PATH", "/Shared"));
        assertThat(config.driveIds()).containsExactly("d1", "d2");
        assertThat(config.folderPath()).isEqualTo("/Shared");
        assertThat(config.hasDriveAllowlist()).isTrue();
        assertThat(valid().driveIds(List.of("x")).build().hasDriveAllowlist()).isTrue();
    }
}
