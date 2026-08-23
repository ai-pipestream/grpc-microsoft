package ai.pipestream.microsoft;

import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Everything the Microsoft Graph crawler needs from the outside world.
 * Production uses {@link #fromEnvironment()}; tests use {@link #builder()}.
 * The client secret is never included in {@link #toString()}.
 */
public record MicrosoftConnectorConfig(
        String tenantId,
        String clientId,
        String clientSecret,
        String siteId,
        List<String> driveIds,
        String folderPath,
        String graphBaseUrl,
        String authority) {

    public static final String ENV_TENANT_ID = "MICROSOFT_TENANT_ID";
    public static final String ENV_CLIENT_ID = "MICROSOFT_CLIENT_ID";
    public static final String ENV_CLIENT_SECRET = "MICROSOFT_CLIENT_SECRET";
    public static final String ENV_SITE_ID = "MICROSOFT_SITE_ID";
    public static final String ENV_DRIVE_IDS = "MICROSOFT_DRIVE_IDS";
    public static final String ENV_FOLDER_PATH = "MICROSOFT_FOLDER_PATH";
    public static final String ENV_GRAPH_BASE_URL = "MICROSOFT_GRAPH_BASE_URL";
    public static final String ENV_AUTHORITY = "MICROSOFT_AUTHORITY";

    public MicrosoftConnectorConfig {
        if (tenantId == null || tenantId.isBlank()) {
            throw new IllegalArgumentException(ENV_TENANT_ID + " is required");
        }
        tenantId = tenantId.trim();
        if (clientId == null || clientId.isBlank()) {
            throw new IllegalArgumentException(ENV_CLIENT_ID + " is required");
        }
        clientId = clientId.trim();
        if (clientSecret == null || clientSecret.isBlank()) {
            throw new IllegalArgumentException(ENV_CLIENT_SECRET + " is required");
        }
        if (siteId == null) {
            siteId = "";
        } else {
            siteId = siteId.trim();
        }
        if (driveIds == null) {
            driveIds = List.of();
        } else {
            driveIds = List.copyOf(driveIds);
        }
        if (folderPath == null || folderPath.isBlank()) {
            folderPath = "/";
        }
        if (graphBaseUrl == null || graphBaseUrl.isBlank()) {
            graphBaseUrl = "https://graph.microsoft.com/v1.0";
        }
        graphBaseUrl = graphBaseUrl.trim().replaceAll("/+$", "");
        if (authority == null || authority.isBlank()) {
            authority = "https://login.microsoftonline.com";
        }
        authority = authority.trim().replaceAll("/+$", "");
    }

    public boolean hasDriveAllowlist() {
        return !driveIds.isEmpty();
    }

    public GraphAuth.Config authConfig() {
        return new GraphAuth.Config(authority, tenantId, clientId, clientSecret);
    }

    public static MicrosoftConnectorConfig fromEnvironment() {
        return fromEnvironment(System.getenv());
    }

    static MicrosoftConnectorConfig fromEnvironment(Map<String, String> env) {
        return builder()
                .tenantId(env.get(ENV_TENANT_ID))
                .clientId(env.get(ENV_CLIENT_ID))
                .clientSecret(env.get(ENV_CLIENT_SECRET))
                .siteId(env.get(ENV_SITE_ID))
                .driveIds(parseList(env.get(ENV_DRIVE_IDS)))
                .folderPath(env.get(ENV_FOLDER_PATH))
                .graphBaseUrl(env.get(ENV_GRAPH_BASE_URL))
                .authority(env.get(ENV_AUTHORITY))
                .build();
    }

    public static Builder builder() {
        return new Builder();
    }

    private static List<String> parseList(String value) {
        if (value == null || value.isBlank()) {
            return List.of();
        }
        return java.util.Arrays.stream(value.split(","))
                .map(String::trim)
                .filter(s -> !s.isEmpty())
                .toList();
    }

    @Override
    public String toString() {
        return "MicrosoftConnectorConfig{tenantId=" + tenantId
                + ", clientId=" + clientId
                + ", clientSecret=***"
                + ", siteId=" + siteId
                + ", driveIds=" + driveIds
                + ", folderPath=" + folderPath
                + ", graphBaseUrl=" + graphBaseUrl
                + ", authority=" + authority + "}";
    }

    public static final class Builder {
        private String tenantId;
        private String clientId;
        private String clientSecret;
        private String siteId = "";
        private List<String> driveIds = List.of();
        private String folderPath = "/";
        private String graphBaseUrl = "https://graph.microsoft.com/v1.0";
        private String authority = "https://login.microsoftonline.com";

        private Builder() {
        }

        public Builder tenantId(String tenantId) {
            this.tenantId = tenantId;
            return this;
        }

        public Builder clientId(String clientId) {
            this.clientId = clientId;
            return this;
        }

        public Builder clientSecret(String clientSecret) {
            this.clientSecret = clientSecret;
            return this;
        }

        public Builder siteId(String siteId) {
            this.siteId = Objects.requireNonNullElse(siteId, "");
            return this;
        }

        public Builder driveIds(List<String> driveIds) {
            this.driveIds = Objects.requireNonNullElse(driveIds, List.of());
            return this;
        }

        public Builder folderPath(String folderPath) {
            this.folderPath = folderPath;
            return this;
        }

        public Builder graphBaseUrl(String graphBaseUrl) {
            this.graphBaseUrl = graphBaseUrl;
            return this;
        }

        public Builder authority(String authority) {
            this.authority = authority;
            return this;
        }

        public MicrosoftConnectorConfig build() {
            return new MicrosoftConnectorConfig(tenantId, clientId, clientSecret, siteId, driveIds,
                    folderPath, graphBaseUrl, authority);
        }
    }
}
