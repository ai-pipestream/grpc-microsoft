package ai.pipestream.microsoft.connector;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/**
 * GCA {@code CustomConfiguration.configuration} JSON decoded by this
 * adapter. The Graph Connector Agent does not interpret the string; we
 * do. Credentials for Microsoft Graph live on {@code MicrosoftService},
 * not here.
 *
 * <pre>
 * {
 *   "target": "localhost:9096",
 *   "plaintext": true,
 *   "driveIds": [],
 *   "folderPath": "/",
 *   "includeContent": false
 * }
 * </pre>
 */
public record ConnectorCustomConfig(
        String target,
        boolean plaintext,
        List<String> driveIds,
        String folderPath,
        boolean includeContent) {

    public static final String ENV_TARGET = "MICROSOFT_GRPC_TARGET";
    public static final String ENV_PLAINTEXT = "MICROSOFT_GRPC_PLAINTEXT";
    public static final String DEFAULT_TARGET = "localhost:9096";

    private static final ObjectMapper JSON = new ObjectMapper();

    public ConnectorCustomConfig {
        if (target == null || target.isBlank()) {
            throw new IllegalArgumentException("target is required (gRPC host:port of MicrosoftService)");
        }
        target = target.trim();
        if (driveIds == null) {
            driveIds = List.of();
        } else {
            driveIds = List.copyOf(driveIds);
        }
        if (folderPath == null || folderPath.isBlank()) {
            folderPath = "/";
        }
    }

    public static ConnectorCustomConfig defaults() {
        return parse("", null);
    }

    /**
     * Decode the admin-supplied JSON, falling back to process environment
     * and {@link #DEFAULT_TARGET}. Blank JSON is valid and means defaults.
     *
     * @param configuration GCA custom-configuration string, possibly blank
     * @param datasourceUrl unused as a gRPC target (GCA's datasource URL
     *        is the SharePoint / Graph location); accepted so callers can
     *        pass the AuthenticationData field through without a second
     *        parse
     */
    public static ConnectorCustomConfig parse(String configuration, String datasourceUrl) {
        String target = firstNonBlank(System.getenv(ENV_TARGET), DEFAULT_TARGET);
        boolean plaintext = parseBoolean(System.getenv(ENV_PLAINTEXT), true);
        List<String> driveIds = List.of();
        String folderPath = "/";
        boolean includeContent = false;
        String json = configuration == null ? "" : configuration.trim();
        if (!json.isEmpty()) {
            JsonNode node;
            try {
                node = JSON.readTree(json);
            } catch (IOException e) {
                throw new IllegalArgumentException("custom configuration is not JSON: " + e.getMessage(), e);
            }
            if (node == null || !node.isObject()) {
                throw new IllegalArgumentException("custom configuration must be a JSON object");
            }
            target = firstNonBlank(text(node, "target"), target);
            if (node.has("plaintext")) {
                plaintext = node.path("plaintext").asBoolean(plaintext);
            }
            if (node.path("driveIds").isArray()) {
                List<String> ids = new ArrayList<>();
                node.path("driveIds").forEach(item -> {
                    if (item.isTextual() && !item.asText().isBlank()) {
                        ids.add(item.asText().trim());
                    }
                });
                driveIds = ids;
            }
            folderPath = firstNonBlank(text(node, "folderPath"), folderPath);
            if (node.has("includeContent")) {
                includeContent = node.path("includeContent").asBoolean(false);
            }
        }
        String fromDatasource = grpcTarget(datasourceUrl);
        if (!fromDatasource.isEmpty() && (target.equals(DEFAULT_TARGET) || json.isEmpty())) {
            // GCA's datasource URL is usually a SharePoint site. When an
            // admin puts host:port (or grpc://host:port) there instead, it
            // is the MicrosoftService target — the only place
            // ValidateAuthentication can send one.
            target = fromDatasource;
        }
        Objects.requireNonNullElse(datasourceUrl, "");
        return new ConnectorCustomConfig(target, plaintext, driveIds, folderPath, includeContent);
    }

    static String grpcTarget(String datasourceUrl) {
        if (datasourceUrl == null || datasourceUrl.isBlank()) {
            return "";
        }
        String value = datasourceUrl.trim();
        if (value.startsWith("grpc://")) {
            return value.substring("grpc://".length());
        }
        if (!value.contains("://") && value.contains(":")) {
            return value;
        }
        return "";
    }

    private static String text(JsonNode node, String field) {
        JsonNode value = node.path(field);
        return value.isTextual() ? value.asText() : "";
    }

    private static String firstNonBlank(String... values) {
        for (String value : values) {
            if (value != null && !value.isBlank()) {
                return value;
            }
        }
        return "";
    }

    private static boolean parseBoolean(String value, boolean fallback) {
        if (value == null || value.isBlank()) {
            return fallback;
        }
        return switch (value.trim().toLowerCase(Locale.ROOT)) {
            case "1", "true", "yes" -> true;
            case "0", "false", "no" -> false;
            default -> fallback;
        };
    }
}
