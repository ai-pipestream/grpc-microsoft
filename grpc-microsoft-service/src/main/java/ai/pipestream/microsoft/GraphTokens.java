package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.util.Base64;

/**
 * Reads the auth mode out of an Entra access token's claims. The token is not
 * verified (Graph does that); only the {@code idtyp}, {@code scp} and
 * {@code roles} claims are read, and nothing from the token is logged or
 * returned.
 */
final class GraphTokens {

    private static final ObjectMapper JSON = new ObjectMapper();

    private GraphTokens() {
    }

    /**
     * The auth mode a JWT access token declares.
     *
     * @param token the bearer token
     * @return application for {@code idtyp=app} or a {@code roles}-only token,
     *         delegated for a token with {@code scp}, unspecified otherwise
     */
    static AuthMode authMode(String token) {
        if (token == null) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
        int first = token.indexOf('.');
        int second = first < 0 ? -1 : token.indexOf('.', first + 1);
        if (first < 0 || second < 0) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
        JsonNode claims;
        try {
            claims = JSON.readTree(Base64.getUrlDecoder().decode(token.substring(first + 1, second)));
        } catch (IOException | IllegalArgumentException e) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
        if (claims == null || !claims.isObject()) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
        if ("app".equalsIgnoreCase(claims.path("idtyp").asText(""))) {
            return AuthMode.AUTH_MODE_APPLICATION;
        }
        if (claims.hasNonNull("scp")) {
            return AuthMode.AUTH_MODE_DELEGATED;
        }
        if (claims.hasNonNull("roles")) {
            return AuthMode.AUTH_MODE_APPLICATION;
        }
        return AuthMode.AUTH_MODE_UNSPECIFIED;
    }
}
