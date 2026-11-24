package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.security.KeyFactory;
import java.security.MessageDigest;
import java.security.PrivateKey;
import java.security.Signature;
import java.security.cert.CertificateFactory;
import java.security.spec.PKCS8EncodedKeySpec;
import java.time.Instant;
import java.util.Base64;
import java.util.UUID;

/**
 * An app registration's certificate credential: the RSA private key and the
 * certificate uploaded to Entra. Signs the client assertion that replaces a
 * client secret in the client-credentials flow.
 */
public final class ClientCertificate {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final Base64.Encoder URL = Base64.getUrlEncoder().withoutPadding();
    /** DER of AlgorithmIdentifier { rsaEncryption, NULL }. */
    private static final byte[] RSA_ALGORITHM = {
        0x30, 0x0d, 0x06, 0x09, 0x2a, (byte) 0x86, 0x48, (byte) 0x86, (byte) 0xf7, 0x0d, 0x01, 0x01, 0x01,
        0x05, 0x00};
    /** How long an assertion is valid. */
    private static final long LIFETIME_SECONDS = 600;

    private final PrivateKey key;
    private final String thumbprint;

    private ClientCertificate(PrivateKey key, String thumbprint) {
        this.key = key;
        this.thumbprint = thumbprint;
    }

    /**
     * Reads a PEM text holding one RSA private key ({@code PRIVATE KEY} or
     * {@code RSA PRIVATE KEY}) and one {@code CERTIFICATE}, in any order.
     *
     * @param pem the PEM text
     * @return the credential
     * @throws IllegalArgumentException if a block is missing or unreadable; the
     *         message never quotes the PEM
     */
    public static ClientCertificate fromPem(String pem) {
        byte[] pkcs8 = block(pem, "PRIVATE KEY");
        if (pkcs8 == null) {
            byte[] pkcs1 = block(pem, "RSA PRIVATE KEY");
            if (pkcs1 == null) {
                throw new IllegalArgumentException("the certificate PEM holds no private key block");
            }
            pkcs8 = wrapPkcs1(pkcs1);
        }
        byte[] certificate = block(pem, "CERTIFICATE");
        if (certificate == null) {
            throw new IllegalArgumentException("the certificate PEM holds no certificate block");
        }
        try {
            PrivateKey key = KeyFactory.getInstance("RSA").generatePrivate(new PKCS8EncodedKeySpec(pkcs8));
            byte[] der = CertificateFactory.getInstance("X.509")
                    .generateCertificate(new ByteArrayInputStream(certificate)).getEncoded();
            return new ClientCertificate(key,
                    URL.encodeToString(MessageDigest.getInstance("SHA-1").digest(der)));
        } catch (GeneralSecurityException e) {
            throw new IllegalArgumentException("the certificate PEM could not be read: "
                    + e.getClass().getSimpleName(), e);
        }
    }

    /**
     * A signed client assertion (RS256 JWT) for the token endpoint.
     *
     * @param clientId application (client) id, the issuer and subject
     * @param audience the token endpoint URL
     * @param now issue time
     * @return the compact JWT
     */
    public String assertion(String clientId, String audience, Instant now) {
        ObjectNode header = JSON.createObjectNode();
        header.put("alg", "RS256");
        header.put("typ", "JWT");
        header.put("x5t", thumbprint);
        ObjectNode claims = JSON.createObjectNode();
        claims.put("aud", audience);
        claims.put("iss", clientId);
        claims.put("sub", clientId);
        claims.put("jti", UUID.randomUUID().toString());
        claims.put("nbf", now.getEpochSecond());
        claims.put("exp", now.getEpochSecond() + LIFETIME_SECONDS);
        String signed = encode(header) + "." + encode(claims);
        try {
            Signature signature = Signature.getInstance("SHA256withRSA");
            signature.initSign(key);
            signature.update(signed.getBytes(StandardCharsets.US_ASCII));
            return signed + "." + URL.encodeToString(signature.sign());
        } catch (GeneralSecurityException e) {
            throw new IllegalStateException("the client assertion could not be signed", e);
        }
    }

    /** The base64url SHA-1 thumbprint of the certificate ({@code x5t}). */
    String thumbprint() {
        return thumbprint;
    }

    @Override
    public String toString() {
        return "ClientCertificate{x5t=" + thumbprint + ", key=***}";
    }

    private static String encode(ObjectNode node) {
        return URL.encodeToString(node.toString().getBytes(StandardCharsets.UTF_8));
    }

    /** The decoded body of the first PEM block with this label, or null. */
    private static byte[] block(String pem, String label) {
        String begin = "-----BEGIN " + label + "-----";
        String end = "-----END " + label + "-----";
        int from = pem.indexOf(begin);
        int to = from < 0 ? -1 : pem.indexOf(end, from);
        if (to < 0) {
            return null;
        }
        try {
            return Base64.getMimeDecoder().decode(pem.substring(from + begin.length(), to));
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("the " + label + " block is not base64", e);
        }
    }

    /** Wraps a PKCS#1 RSA key in a PKCS#8 PrivateKeyInfo. */
    private static byte[] wrapPkcs1(byte[] pkcs1) {
        ByteArrayOutputStream body = new ByteArrayOutputStream();
        body.writeBytes(new byte[] {0x02, 0x01, 0x00});
        body.writeBytes(RSA_ALGORITHM);
        body.writeBytes(der(0x04, pkcs1));
        return der(0x30, body.toByteArray());
    }

    private static byte[] der(int tag, byte[] content) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        out.write(tag);
        int length = content.length;
        if (length < 0x80) {
            out.write(length);
        } else {
            int bytes = (32 - Integer.numberOfLeadingZeros(length) + 7) / 8;
            out.write(0x80 | bytes);
            for (int i = bytes - 1; i >= 0; i--) {
                out.write(length >>> (8 * i));
            }
        }
        out.writeBytes(content);
        return out.toByteArray();
    }
}
