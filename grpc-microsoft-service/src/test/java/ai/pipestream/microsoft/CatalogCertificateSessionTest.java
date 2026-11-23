package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.sync.v1.Connection;
import ai.pipestream.sync.v1.ConnectionKind;
import ai.pipestream.sync.v1.ConnectionServiceGrpc;
import ai.pipestream.sync.v1.GetConnectionRequest;
import ai.pipestream.sync.v1.GetConnectionResponse;
import com.sun.net.httpserver.HttpServer;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import java.io.InputStream;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CopyOnWriteArrayList;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** A catalog or inline connection with a certificate signs in with it and reaches Graph. */
class CatalogCertificateSessionTest {

    private final Map<String, Connection> rows = new ConcurrentHashMap<>();
    private final List<String> tokenRequests = new CopyOnWriteArrayList<>();
    private final List<String> graphRequests = new CopyOnWriteArrayList<>();
    private HttpServer http;
    private Server catalog;
    private ManagedChannel channel;
    private MicrosoftSessions sessions;

    @BeforeEach
    void start() throws Exception {
        http = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        http.createContext("/", exchange -> {
            String path = exchange.getRequestURI().getPath();
            String sent = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
            String answer;
            if (path.endsWith("/oauth2/v2.0/token")) {
                tokenRequests.add(path + " " + sent);
                answer = "{\"access_token\":\"catalog-token\",\"expires_in\":3600}";
            } else {
                graphRequests.add(path + " " + exchange.getRequestHeaders().getFirst("authorization"));
                answer = "{\"id\":\"site-1\"}";
            }
            byte[] bytes = answer.getBytes(StandardCharsets.UTF_8);
            exchange.getResponseHeaders().add("content-type", "application/json");
            exchange.sendResponseHeaders(200, bytes.length);
            exchange.getResponseBody().write(bytes);
            exchange.close();
        });
        http.start();
        String base = "http://127.0.0.1:" + http.getAddress().getPort();
        String name = InProcessServerBuilder.generateName();
        catalog = InProcessServerBuilder.forName(name).directExecutor()
                .addService(new ConnectionServiceGrpc.ConnectionServiceImplBase() {
                    @Override
                    public void getConnection(GetConnectionRequest request,
                            StreamObserver<GetConnectionResponse> observer) {
                        Connection row = rows.get(request.getConnectionId());
                        if (row == null || !request.getIncludeSecret()) {
                            observer.onError(Status.NOT_FOUND.asRuntimeException());
                            return;
                        }
                        observer.onNext(GetConnectionResponse.newBuilder().setConnection(row).build());
                        observer.onCompleted();
                    }
                }).build().start();
        channel = InProcessChannelBuilder.forName(name).directExecutor().build();
        MicrosoftConnectorConfig defaults = MicrosoftConnectorConfig.builder().tenantId("default-tenant")
                .clientId("default-client").clientSecret("default-secret").graphBaseUrl(base)
                .authority(base).build();
        sessions = new MicrosoftSessions(defaults, new GraphFiles(new GraphClient(base, () -> "default-token")),
                ConnectionServiceGrpc.newBlockingStub(channel));
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        catalog.shutdownNow();
        http.stop(0);
    }

    private static String pem() throws Exception {
        try (InputStream in = CatalogCertificateSessionTest.class.getResourceAsStream("/certs/test-only-pkcs8.pem")) {
            return new String(in.readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    private static Connection.Builder row(String id) {
        return Connection.newBuilder().setConnectionId(id).setKind(ConnectionKind.CONNECTION_KIND_MICROSOFT)
                .setTenantId("contoso-tenant").setClientId("contoso-client");
    }

    @Test
    void aCatalogConnectionWithOnlyACertificateSignsInWithIt() throws Exception {
        rows.put("contoso", row("contoso").setClientCertificatePem(pem()).build());

        MicrosoftSessions.Session session = sessions.resolve("contoso");
        assertThat(session.graph().get("/sites/root").path("id").asText()).isEqualTo("site-1");

        assertThat(tokenRequests).singleElement().satisfies(request -> assertThat(request)
                .startsWith("/contoso-tenant/oauth2/v2.0/token ")
                .contains("client_id=contoso-client", "client_assertion=", "client_assertion_type=")
                .doesNotContain("client_secret"));
        assertThat(graphRequests).containsExactly("/sites/root Bearer catalog-token");
        assertThat(session.config().toString()).doesNotContain("BEGIN");
    }

    @Test
    void anInlineCertificateSignsIn() throws Exception {
        MicrosoftSessions.Session session = sessions.inline("contoso-tenant", "contoso-client", "", pem());

        assertThat(session.graph().get("/sites/root").path("id").asText()).isEqualTo("site-1");
        assertThat(tokenRequests).singleElement().satisfies(request -> assertThat(request)
                .contains("client_assertion=").doesNotContain("client_secret"));
    }

    @Test
    void aCatalogConnectionWithNoCredentialIsRefused() {
        rows.put("empty", row("empty").build());

        assertThatThrownBy(() -> sessions.resolve("empty"))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(
                        e.getStatus().getCode()).isEqualTo(Status.Code.FAILED_PRECONDITION));
        assertThat(tokenRequests).isEmpty();
    }
}
