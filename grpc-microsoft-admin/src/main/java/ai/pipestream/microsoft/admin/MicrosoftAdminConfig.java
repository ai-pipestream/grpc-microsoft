package ai.pipestream.microsoft.admin;

import ai.pipestream.microsoft.ClientCertificate;
import ai.pipestream.microsoft.GraphAuth;
import ai.pipestream.microsoft.GraphLinks;
import java.util.Map;

/**
 * What the admin service needs from the environment. The client secret and
 * the certificate are never included in {@link #toString()}.
 *
 * @param tenantId Microsoft Entra tenant id ({@code MICROSOFT_TENANT_ID}); required
 * @param clientId application (client) id of the admin identity
 *        ({@code MICROSOFT_ADMIN_CLIENT_ID}); required
 * @param clientSecret its client secret ({@code MICROSOFT_ADMIN_CLIENT_SECRET});
 *        required unless a certificate is set
 * @param clientCertificatePem its private key and certificate in PEM
 *        ({@code MICROSOFT_ADMIN_CLIENT_CERT_PEM}); empty = sign in with the secret
 * @param graphBaseUrl Microsoft Graph base URL ({@code MICROSOFT_GRAPH_BASE_URL},
 *        default {@code https://graph.microsoft.com/v1.0})
 * @param authority Entra authority host ({@code MICROSOFT_AUTHORITY},
 *        default {@code https://login.microsoftonline.com})
 */
public record MicrosoftAdminConfig(String tenantId, String clientId, String clientSecret,
        String clientCertificatePem, String graphBaseUrl, String authority) {

    /** Environment variable for the Microsoft Entra tenant id. */
    public static final String ENV_TENANT_ID = "MICROSOFT_TENANT_ID";
    /** Environment variable for the admin identity's application (client) id. */
    public static final String ENV_CLIENT_ID = "MICROSOFT_ADMIN_CLIENT_ID";
    /** Environment variable for the admin identity's client secret. */
    public static final String ENV_CLIENT_SECRET = "MICROSOFT_ADMIN_CLIENT_SECRET";
    /** Environment variable for the admin identity's certificate credential. */
    public static final String ENV_CLIENT_CERT_PEM = "MICROSOFT_ADMIN_CLIENT_CERT_PEM";
    /** Environment variable for the Microsoft Graph base URL. */
    public static final String ENV_GRAPH_BASE_URL = "MICROSOFT_GRAPH_BASE_URL";
    /** Environment variable for the Entra authority host. */
    public static final String ENV_AUTHORITY = "MICROSOFT_AUTHORITY";

    /**
     * Trims and defaults fields; the tenant, the client id and one of the
     * secret or the certificate are required.
     */
    public MicrosoftAdminConfig {
        tenantId = required(tenantId, ENV_TENANT_ID);
        clientId = required(clientId, ENV_CLIENT_ID);
        clientCertificatePem = clientCertificatePem == null || clientCertificatePem.isBlank()
                ? "" : clientCertificatePem;
        if ((clientSecret == null || clientSecret.isBlank()) && clientCertificatePem.isEmpty()) {
            throw new IllegalArgumentException(ENV_CLIENT_SECRET + " or " + ENV_CLIENT_CERT_PEM
                    + " is required");
        }
        graphBaseUrl = GraphLinks.stripTrailingSlashes(graphBaseUrl == null || graphBaseUrl.isBlank()
                ? "https://graph.microsoft.com/v1.0" : graphBaseUrl.trim());
        authority = GraphLinks.stripTrailingSlashes(authority == null || authority.isBlank()
                ? "https://login.microsoftonline.com" : authority.trim());
    }

    private static String required(String value, String name) {
        if (value == null || value.isBlank()) {
            throw new IllegalArgumentException(name + " is required");
        }
        return value.trim();
    }

    /**
     * Builds the config from the process environment.
     *
     * @return the resolved config
     */
    public static MicrosoftAdminConfig fromEnvironment() {
        return fromEnvironment(System.getenv());
    }

    /**
     * Builds the config from an environment map.
     *
     * @param env the environment to read
     * @return the resolved config
     */
    public static MicrosoftAdminConfig fromEnvironment(Map<String, String> env) {
        return new MicrosoftAdminConfig(env.get(ENV_TENANT_ID), env.get(ENV_CLIENT_ID),
                env.get(ENV_CLIENT_SECRET), env.get(ENV_CLIENT_CERT_PEM), env.get(ENV_GRAPH_BASE_URL),
                env.get(ENV_AUTHORITY));
    }

    /**
     * The Entra client-credentials config of the admin identity.
     *
     * @return authority, tenant, client id, client secret and certificate
     */
    public GraphAuth.Config authConfig() {
        return new GraphAuth.Config(authority, tenantId, clientId, clientSecret,
                clientCertificatePem.isEmpty() ? null : ClientCertificate.fromPem(clientCertificatePem));
    }

    /** Secrets stay out of any log line this record lands in. */
    @Override
    public String toString() {
        return "MicrosoftAdminConfig{tenantId=" + tenantId + ", clientId=" + clientId
                + ", clientSecret=***, clientCertificatePem="
                + (clientCertificatePem.isEmpty() ? "" : "***") + ", graphBaseUrl=" + graphBaseUrl
                + ", authority=" + authority + "}";
    }
}
