package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.sun.net.httpserver.HttpServer;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.net.InetSocketAddress;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.Signature;
import java.security.cert.Certificate;
import java.security.cert.CertificateFactory;
import java.time.Instant;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;
import org.junit.jupiter.api.Test;

class ClientCertificateTest {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static String pem(String name) throws Exception {
        try (InputStream in = ClientCertificateTest.class.getResourceAsStream("/certs/" + name)) {
            return new String(in.readAllBytes(), StandardCharsets.US_ASCII);
        }
    }

    private static Certificate certificate(String pem) throws Exception {
        String end = "-----END CERTIFICATE-----";
        String block = pem.substring(pem.indexOf("-----BEGIN CERTIFICATE-----"), pem.indexOf(end) + end.length());
        return CertificateFactory.getInstance("X.509")
                .generateCertificate(new ByteArrayInputStream(block.getBytes(StandardCharsets.US_ASCII)));
    }

    private static void verify(String jwt, Certificate certificate) throws Exception {
        int last = jwt.lastIndexOf('.');
        Signature signature = Signature.getInstance("SHA256withRSA");
        signature.initVerify(certificate.getPublicKey());
        signature.update(jwt.substring(0, last).getBytes(StandardCharsets.US_ASCII));
        assertThat(signature.verify(Base64.getUrlDecoder().decode(jwt.substring(last + 1)))).isTrue();
    }

    private static JsonNode part(String jwt, int index) throws Exception {
        List<String> parts = List.of(jwt.split("\\."));
        assertThat(parts).hasSize(3);
        return JSON.readTree(Base64.getUrlDecoder().decode(parts.get(index)));
    }

    @Test
    void theAssertionIsSignedByTheCertificateAndNamesIt() throws Exception {
        String pem = pem("test-only-pkcs8.pem");
        Certificate certificate = certificate(pem);
        Instant now = Instant.ofEpochSecond(1_800_000_000L);

        String jwt = ClientCertificate.fromPem(pem).assertion("client-1", "https://login.example/t/token", now);

        verify(jwt, certificate);
        JsonNode header = part(jwt, 0);
        assertThat(header.path("alg").asText()).isEqualTo("RS256");
        assertThat(header.path("typ").asText()).isEqualTo("JWT");
        assertThat(header.path("x5t").asText()).isEqualTo(Base64.getUrlEncoder().withoutPadding()
                .encodeToString(MessageDigest.getInstance("SHA-1").digest(certificate.getEncoded())));
        JsonNode claims = part(jwt, 1);
        assertThat(claims.path("aud").asText()).isEqualTo("https://login.example/t/token");
        assertThat(claims.path("iss").asText()).isEqualTo("client-1");
        assertThat(claims.path("sub").asText()).isEqualTo("client-1");
        assertThat(claims.path("jti").asText()).isNotEmpty();
        assertThat(claims.path("nbf").asLong()).isEqualTo(1_800_000_000L);
        assertThat(claims.path("exp").asLong()).isEqualTo(1_800_000_600L);
    }

    @Test
    void aTraditionalRsaKeyBlockIsRead() throws Exception {
        String pem = pem("test-only-pkcs1.pem");
        ClientCertificate credential = ClientCertificate.fromPem(pem);

        verify(credential.assertion("client-1", "aud", Instant.now()), certificate(pem));
        assertThat(credential.thumbprint())
                .isEqualTo(ClientCertificate.fromPem(pem("test-only-pkcs8.pem")).thumbprint());
    }

    @Test
    void aMissingBlockIsRefusedWithoutQuotingThePem() throws Exception {
        String pem = pem("test-only-pkcs8.pem");
        String keyOnly = pem.substring(0, pem.indexOf("-----BEGIN CERTIFICATE-----"));
        String certificateOnly = pem.substring(pem.indexOf("-----BEGIN CERTIFICATE-----"));

        assertThatThrownBy(() -> ClientCertificate.fromPem(keyOnly))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessage("the certificate PEM holds no certificate block");
        assertThatThrownBy(() -> ClientCertificate.fromPem(certificateOnly))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessage("the certificate PEM holds no private key block");
        assertThat(ClientCertificate.fromPem(pem).toString()).doesNotContain("MII");
    }

    @Test
    void theTokenRequestCarriesTheAssertionAndNoSecret() throws Exception {
        String pem = pem("test-only-pkcs8.pem");
        AtomicReference<Map<String, String>> seen = new AtomicReference<>();
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/", exchange -> {
            Map<String, String> form = new LinkedHashMap<>();
            String body = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
            for (String pair : body.split("&")) {
                int eq = pair.indexOf('=');
                form.put(URLDecoder.decode(pair.substring(0, eq), StandardCharsets.UTF_8),
                        URLDecoder.decode(pair.substring(eq + 1), StandardCharsets.UTF_8));
            }
            seen.set(form);
            byte[] answer = "{\"access_token\":\"t-1\",\"expires_in\":3600}".getBytes(StandardCharsets.UTF_8);
            exchange.getResponseHeaders().add("content-type", "application/json");
            exchange.sendResponseHeaders(200, answer.length);
            exchange.getResponseBody().write(answer);
            exchange.close();
        });
        server.start();
        try {
            String authority = "http://127.0.0.1:" + server.getAddress().getPort();
            MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder().tenantId("tenant")
                    .clientId("client-1").clientSecret("unused-secret").clientCertificatePem(pem)
                    .authority(authority).build();

            GraphAuth.Token token = new GraphAuth(config.authConfig())
                    .clientCredentials("https://contoso.sharepoint.com/.default");

            assertThat(token.accessToken()).isEqualTo("t-1");
            Map<String, String> form = seen.get();
            assertThat(form).containsEntry("grant_type", "client_credentials")
                    .containsEntry("client_id", "client-1")
                    .containsEntry("scope", "https://contoso.sharepoint.com/.default")
                    .containsEntry("client_assertion_type",
                            "urn:ietf:params:oauth:client-assertion-type:jwt-bearer")
                    .doesNotContainKey("client_secret");
            String jwt = form.get("client_assertion");
            verify(jwt, certificate(pem));
            assertThat(part(jwt, 1).path("aud").asText())
                    .isEqualTo(authority + "/tenant/oauth2/v2.0/token");
            assertThat(config.toString()).doesNotContain("BEGIN").contains("clientCertificatePem=***");
        } finally {
            server.stop(0);
        }
    }
}
