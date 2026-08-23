package ai.pipestream.microsoft;

import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Everything the Microsoft Graph crawler needs from the outside world, in one
 * value object - no framework configuration binding. Production uses
 * {@link #fromEnvironment()}; tests use {@link #builder()}. The client secret
 * is never included in {@link #toString()}.
 *
 * @param tenantId Microsoft Entra tenant id ({@code MICROSOFT_TENANT_ID});
 *        required
 * @param clientId application (client) id ({@code MICROSOFT_CLIENT_ID});
 *        required
 * @param clientSecret client secret ({@code MICROSOFT_CLIENT_SECRET});
 *        required, redacted everywhere
 * @param siteId SharePoint site id ({@code MICROSOFT_SITE_ID}); empty skips
 *        site-scoped listing
 * @param driveIds drive ids to crawl ({@code MICROSOFT_DRIVE_IDS},
 *        comma-separated); empty = every drive the credentials can see
 * @param folderPath folder path within each drive ({@code MICROSOFT_FOLDER_PATH},
 *        default {@code "/"})
 * @param graphBaseUrl Microsoft Graph base URL ({@code MICROSOFT_GRAPH_BASE_URL},
 *        default {@code https://graph.microsoft.com/v1.0})
 * @param authority Entra authority host ({@code MICROSOFT_AUTHORITY},
 *        default {@code https://login.microsoftonline.com})
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

    /** Environment variable for the Microsoft Entra tenant id. */
    public static final String ENV_TENANT_ID = "MICROSOFT_TENANT_ID";
    /** Environment variable for the application (client) id. */
    public static final String ENV_CLIENT_ID = "MICROSOFT_CLIENT_ID";
    /** Environment variable for the client secret. */
    public static final String ENV_CLIENT_SECRET = "MICROSOFT_CLIENT_SECRET";
    /** Environment variable for the SharePoint site id. */
    public static final String ENV_SITE_ID = "MICROSOFT_SITE_ID";
    /** Environment variable for the drive-id allowlist (comma-separated). */
    public static final String ENV_DRIVE_IDS = "MICROSOFT_DRIVE_IDS";
    /** Environment variable for the folder path within each drive. */
    public static final String ENV_FOLDER_PATH = "MICROSOFT_FOLDER_PATH";
    /** Environment variable for the Microsoft Graph base URL. */
    public static final String ENV_GRAPH_BASE_URL = "MICROSOFT_GRAPH_BASE_URL";
    /** Environment variable for the Entra authority host. */
    public static final String ENV_AUTHORITY = "MICROSOFT_AUTHORITY";

    /**
     * Trims and defaults fields. {@code tenantId}, {@code clientId}, and
     * {@code clientSecret} are required; blank {@code siteId} stays empty; a
     * null drive list becomes empty; blank {@code folderPath} becomes
     * {@code "/"}; blank Graph and authority URLs take the public-cloud
     * defaults.
     */
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

    /**
     * Whether the crawl is restricted to a drive-id allowlist.
     *
     * @return {@code true} when specific drive ids are configured
     */
    public boolean hasDriveAllowlist() {
        return !driveIds.isEmpty();
    }

    /**
     * The Entra client-credentials config derived from this record.
     *
     * @return authority, tenant, client id, and client secret
     */
    public GraphAuth.Config authConfig() {
        return new GraphAuth.Config(authority, tenantId, clientId, clientSecret);
    }

    /**
     * Build the config from the process environment, using the
     * {@code MICROSOFT_*} variables documented on this record.
     *
     * @return the resolved config
     */
    public static MicrosoftConnectorConfig fromEnvironment() {
        return fromEnvironment(System.getenv());
    }

    /**
     * Build the config from an explicit environment map; production calls
     * {@link #fromEnvironment()}, tests call this.
     *
     * @param env the environment to read
     * @return the resolved config
     */
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

    /**
     * A builder with public-cloud defaults; required credentials are still
     * unset until the corresponding setters run.
     *
     * @return a new builder
     */
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

    /** Secrets stay out of any log line this record lands in. */
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

    /** Test-friendly builder; every field the record validates is optional here. */
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

        /**
         * Sets the Entra tenant id.
         *
         * @param tenantId tenant id
         * @return this builder
         */
        public Builder tenantId(String tenantId) {
            this.tenantId = tenantId;
            return this;
        }

        /**
         * Sets the application (client) id.
         *
         * @param clientId client id
         * @return this builder
         */
        public Builder clientId(String clientId) {
            this.clientId = clientId;
            return this;
        }

        /**
         * Sets the client secret.
         *
         * @param clientSecret client secret
         * @return this builder
         */
        public Builder clientSecret(String clientSecret) {
            this.clientSecret = clientSecret;
            return this;
        }

        /**
         * Sets the SharePoint site id; {@code null} becomes empty.
         *
         * @param siteId site id
         * @return this builder
         */
        public Builder siteId(String siteId) {
            this.siteId = Objects.requireNonNullElse(siteId, "");
            return this;
        }

        /**
         * Sets the drive-id allowlist; {@code null} becomes empty.
         *
         * @param driveIds drive ids
         * @return this builder
         */
        public Builder driveIds(List<String> driveIds) {
            this.driveIds = Objects.requireNonNullElse(driveIds, List.of());
            return this;
        }

        /**
         * Sets the folder path within each drive.
         *
         * @param folderPath folder path
         * @return this builder
         */
        public Builder folderPath(String folderPath) {
            this.folderPath = folderPath;
            return this;
        }

        /**
         * Sets the Microsoft Graph base URL.
         *
         * @param graphBaseUrl Graph root URL
         * @return this builder
         */
        public Builder graphBaseUrl(String graphBaseUrl) {
            this.graphBaseUrl = graphBaseUrl;
            return this;
        }

        /**
         * Sets the Entra authority host.
         *
         * @param authority authority URL
         * @return this builder
         */
        public Builder authority(String authority) {
            this.authority = authority;
            return this;
        }

        /**
         * Builds a config; the record compact constructor validates required
         * fields.
         *
         * @return the config
         */
        public MicrosoftConnectorConfig build() {
            return new MicrosoftConnectorConfig(tenantId, clientId, clientSecret, siteId, driveIds,
                    folderPath, graphBaseUrl, authority);
        }
    }
}
