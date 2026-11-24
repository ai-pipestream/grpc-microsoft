package ai.pipestream.microsoft.admin;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.GraphClient;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.EnsureSiteAppGrantResponse;
import ai.pipestream.microsoft.v1.ListSiteAppGrantsRequest;
import ai.pipestream.microsoft.v1.MicrosoftAdminServiceGrpc;
import ai.pipestream.microsoft.v1.RevokeSiteAppGrantRequest;
import ai.pipestream.microsoft.v1.SiteAppGrant;
import ai.pipestream.microsoft.v1.SiteAppGrantState;
import ai.pipestream.microsoft.v1.SiteAppRole;
import com.sun.net.httpserver.HttpServer;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class MicrosoftAdminGrpcServiceTest {

    private static final String SITE = "contoso.sharepoint.com,s1,w1";
    private static final String PERMISSIONS = "/sites/contoso.sharepoint.com,s1,w1/permissions";
    private static final String APP = "11111111-2222-3333-4444-555555555555";

    /** The site permissions the fake Graph holds: id to roles and application id. */
    private final Map<String, String[]> held = new LinkedHashMap<>();
    private final List<String> requests = new CopyOnWriteArrayList<>();
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private HttpServer graph;
    private Server server;
    private ManagedChannel channel;
    private MicrosoftAdminServiceGrpc.MicrosoftAdminServiceBlockingStub stub;

    @BeforeEach
    void start() throws Exception {
        graph = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        graph.createContext("/", exchange -> {
            String method = exchange.getRequestMethod();
            String path = exchange.getRequestURI().getPath();
            String sent = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
            requests.add(method + " " + path + (sent.isEmpty() ? "" : " " + sent));
            int status = 200;
            String body = "";
            synchronized (held) {
                String id = path.startsWith(PERMISSIONS + "/") ? path.substring(PERMISSIONS.length() + 1) : "";
                if (!"Bearer admin-token".equals(exchange.getRequestHeaders().getFirst("authorization"))) {
                    status = 401;
                    body = "{\"error\":{\"code\":\"InvalidAuthenticationToken\",\"message\":\"no token\"}}";
                } else if (method.equals("GET") && path.equals(PERMISSIONS)) {
                    // Like Graph, the listing leaves the roles out.
                    List<String> rows = new ArrayList<>();
                    held.forEach((key, value) -> rows.add(permission(key, null, value[1])));
                    body = "{\"value\":[" + String.join(",", rows) + "]}";
                } else if (method.equals("POST") && path.equals(PERMISSIONS)) {
                    String newId = "perm-" + (held.size() + 1);
                    held.put(newId, new String[] {between(sent, "\"roles\":[\"", "\""),
                            between(sent, "\"id\":\"", "\"")});
                    status = 201;
                    body = permission(newId, held.get(newId)[0], held.get(newId)[1]);
                } else if (held.containsKey(id) && method.equals("GET")) {
                    body = permission(id, held.get(id)[0], held.get(id)[1]);
                } else if (held.containsKey(id) && method.equals("PATCH")) {
                    held.get(id)[0] = between(sent, "\"roles\":[\"", "\"");
                    body = permission(id, held.get(id)[0], held.get(id)[1]);
                } else if (held.containsKey(id) && method.equals("DELETE")) {
                    held.remove(id);
                    status = 204;
                } else {
                    status = 404;
                    body = "{\"error\":{\"code\":\"itemNotFound\",\"message\":\"not found\"}}";
                }
            }
            byte[] bytes = body.getBytes(StandardCharsets.UTF_8);
            exchange.getResponseHeaders().add("content-type", "application/json");
            exchange.sendResponseHeaders(status, bytes.length == 0 ? -1 : bytes.length);
            if (bytes.length > 0) {
                exchange.getResponseBody().write(bytes);
            }
            exchange.close();
        });
        graph.start();
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftAdminGrpcService(new GraphClient(
                        "http://127.0.0.1:" + graph.getAddress().getPort(), () -> "admin-token")))
                .build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftAdminServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        graph.stop(0);
    }

    private static String permission(String id, String role, String appId) {
        return "{\"id\":\"" + id + "\"," + (role == null ? "" : "\"roles\":[\"" + role + "\"],")
                + "\"grantedToIdentitiesV2\":[{\"application\":{\"id\":\"" + appId
                + "\",\"displayName\":\"app " + appId.substring(0, 4) + "\"}}]}";
    }

    private static String between(String text, String start, String end) {
        int from = text.indexOf(start) + start.length();
        return text.substring(from, text.indexOf(end, from));
    }

    private List<SiteAppGrant> list() {
        List<SiteAppGrant> grants = new ArrayList<>();
        stub.listSiteAppGrants(ListSiteAppGrantsRequest.newBuilder().setSiteId(SITE).build())
                .forEachRemaining(response -> grants.add(response.getGrant()));
        return grants;
    }

    private EnsureSiteAppGrantRequest.Builder ensure(SiteAppRole role) {
        return EnsureSiteAppGrantRequest.newBuilder().setSiteId(SITE).setAppId(APP)
                .setAppDisplayName("connector").setRole(role);
    }

    @Test
    void grantsAreListedWithTheRolesOfEachPermission() {
        held.put("perm-a", new String[] {"manage", APP});
        held.put("perm-b", new String[] {"sharepoint-only-role", "99999999-0000-0000-0000-000000000000"});

        List<SiteAppGrant> grants = list();

        assertThat(grants).hasSize(2);
        assertThat(grants.get(0).getId()).isEqualTo("perm-a");
        assertThat(grants.get(0).getSiteId()).isEqualTo(SITE);
        assertThat(grants.get(0).getAppId()).isEqualTo(APP);
        assertThat(grants.get(0).getAppDisplayName()).isEqualTo("app 1111");
        assertThat(grants.get(0).getRolesList()).containsExactly(SiteAppRole.SITE_APP_ROLE_MANAGE);
        assertThat(grants.get(1).getRolesList()).isEmpty();
        assertThat(grants.get(1).getOtherRolesList()).containsExactly("sharepoint-only-role");
    }

    @Test
    void ensureCreatesThenLeavesAloneThenChangesTheRole() {
        EnsureSiteAppGrantResponse dry = stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_WRITE).setDryRun(true).build());
        assertThat(dry.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_CREATE);
        assertThat(dry.hasGrant()).isFalse();
        assertThat(held).isEmpty();

        EnsureSiteAppGrantResponse created = stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_WRITE).build());
        assertThat(created.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_CREATED);
        assertThat(created.getGrant().getRolesList()).containsExactly(SiteAppRole.SITE_APP_ROLE_WRITE);
        assertThat(created.getGrant().getAppId()).isEqualTo(APP);
        assertThat(requests).contains("POST " + PERMISSIONS + " {\"roles\":[\"write\"],"
                + "\"grantedToIdentities\":[{\"application\":{\"id\":\"" + APP
                + "\",\"displayName\":\"connector\"}}]}");

        requests.clear();
        EnsureSiteAppGrantResponse again = stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_WRITE).build());
        assertThat(again.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_EXISTS);
        assertThat(requests).allMatch(request -> request.startsWith("GET "));

        EnsureSiteAppGrantResponse wouldUpdate = stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_FULL_CONTROL).setDryRun(true).build());
        assertThat(wouldUpdate.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_WOULD_UPDATE);
        assertThat(held.get("perm-1")[0]).isEqualTo("write");

        EnsureSiteAppGrantResponse updated = stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_FULL_CONTROL).build());
        assertThat(updated.getState()).isEqualTo(SiteAppGrantState.SITE_APP_GRANT_STATE_UPDATED);
        assertThat(updated.getPreviousRolesList()).containsExactly(SiteAppRole.SITE_APP_ROLE_WRITE);
        assertThat(updated.getGrant().getRolesList())
                .containsExactly(SiteAppRole.SITE_APP_ROLE_FULL_CONTROL);
        assertThat(held.get("perm-1")[0]).isEqualTo("fullcontrol");
        assertThat(held).hasSize(1);
    }

    @Test
    void revokeRemovesOnlyTheGrantsOfThatApplication() {
        held.put("perm-a", new String[] {"manage", APP});
        held.put("perm-b", new String[] {"read", "99999999-0000-0000-0000-000000000000"});

        assertThat(stub.revokeSiteAppGrant(RevokeSiteAppGrantRequest.newBuilder().setSiteId(SITE)
                .setAppId(APP).build()).getRevoked()).isEqualTo(1);
        assertThat(held).containsOnlyKeys("perm-b");
        assertThat(stub.revokeSiteAppGrant(RevokeSiteAppGrantRequest.newBuilder().setSiteId(SITE)
                .setAppId(APP).build()).getRevoked()).isZero();
    }

    @Test
    void aMissingSiteAppOrRoleIsRefusedBeforeAnyGraphCall() {
        assertThatThrownBy(() -> stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_UNSPECIFIED).build()))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(
                        e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
        assertThatThrownBy(() -> stub.ensureSiteAppGrant(
                ensure(SiteAppRole.SITE_APP_ROLE_READ).setAppId(" ").build()))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(
                        e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
        assertThatThrownBy(() -> list0(""))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(
                        e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
        assertThat(requests).isEmpty();
    }

    private void list0(String siteId) {
        stub.listSiteAppGrants(ListSiteAppGrantsRequest.newBuilder().setSiteId(siteId).build())
                .forEachRemaining(response -> { });
    }

    @Test
    void aConfigNeedsASecretOrACertificateAndHidesBoth() {
        assertThatThrownBy(() -> MicrosoftAdminConfig.fromEnvironment(Map.of(
                MicrosoftAdminConfig.ENV_TENANT_ID, "t", MicrosoftAdminConfig.ENV_CLIENT_ID, "c")))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining(MicrosoftAdminConfig.ENV_CLIENT_CERT_PEM);
        MicrosoftAdminConfig config = MicrosoftAdminConfig.fromEnvironment(Map.of(
                MicrosoftAdminConfig.ENV_TENANT_ID, " t ", MicrosoftAdminConfig.ENV_CLIENT_ID, "c",
                MicrosoftAdminConfig.ENV_CLIENT_SECRET, "hunter2"));
        assertThat(config.tenantId()).isEqualTo("t");
        assertThat(config.graphBaseUrl()).isEqualTo("https://graph.microsoft.com/v1.0");
        assertThat(config.toString()).doesNotContain("hunter2");
    }
}
