package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;

import com.sun.net.httpserver.HttpServer;
import java.net.InetSocketAddress;
import java.net.http.HttpClient;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Base64;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Records every admin scenario against a fake Graph that holds site
 * permissions the way the real one does, then replays the recording.
 */
class AdminGrantRecordingTest {

    private static final String TENANT = "3f2504e0-4f89-41d3-9a0c-0305e82c3301";
    private static final String SITE = "contoso.sharepoint.com,6f9619ff-8b86-4d01-b42d-00cf4fc964ff,"
            + "9b2d5b46-1c2f-4f0e-8a3b-5d7e9c1a2b3c";
    private static final String GRANTEE = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String OTHER = "7c9e6679-7425-40de-944b-e07fc1f90ae7";
    private static final String SELF = "1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed";
    private static final String PERMISSIONS = "/sites/" + SITE + "/permissions";

    /** The site permissions the fake Graph holds: id to roles, application id and display name. */
    private final Map<String, String[]> held = new LinkedHashMap<>();
    private HttpServer graph;

    /** A permission id as Graph forms it: the application and tenant, base64-encoded. */
    private static String permissionId(String appId) {
        return Base64.getUrlEncoder().withoutPadding().encodeToString(
                ("i:0i.t|ms.sp.ext|" + appId + "@" + TENANT).getBytes(StandardCharsets.UTF_8));
    }

    private static String permission(String id, String role, String[] grant) {
        return "{\"id\":\"" + id + "\"," + (role == null ? "" : "\"roles\":[\"" + role + "\"],")
                + "\"grantedToIdentitiesV2\":[{\"application\":{\"id\":\"" + grant[1]
                + "\",\"displayName\":\"" + grant[2] + "\"}}]}";
    }

    private static String between(String text, String start, String end) {
        int from = text.indexOf(start) + start.length();
        return text.substring(from, text.indexOf(end, from));
    }

    @BeforeEach
    void start() throws Exception {
        held.put(permissionId(GRANTEE), new String[] {"manage", GRANTEE, "Contoso Connector"});
        held.put(permissionId(OTHER), new String[] {"read", OTHER, "Contoso Reporting"});
        graph = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        graph.createContext("/", exchange -> {
            String method = exchange.getRequestMethod();
            String path = exchange.getRequestURI().getPath();
            String sent = new String(exchange.getRequestBody().readAllBytes(), StandardCharsets.UTF_8);
            int status = 200;
            String body = "";
            synchronized (held) {
                String id = path.startsWith(PERMISSIONS + "/") ? path.substring(PERMISSIONS.length() + 1) : "";
                if (method.equals("GET") && path.equals(PERMISSIONS)) {
                    // Like Graph, the listing leaves the roles out.
                    List<String> rows = new ArrayList<>();
                    held.forEach((key, value) -> rows.add(permission(key, null, value)));
                    body = "{\"value\":[" + String.join(",", rows) + "]}";
                } else if (method.equals("POST") && path.equals(PERMISSIONS)) {
                    String appId = between(sent, "\"id\":\"", "\"");
                    String[] grant = {between(sent, "\"roles\":[\"", "\""), appId,
                            between(sent, "\"displayName\":\"", "\"")};
                    held.put(permissionId(appId), grant);
                    status = 201;
                    body = permission(permissionId(appId), grant[0], grant);
                } else if (held.containsKey(id) && method.equals("GET")) {
                    body = permission(id, held.get(id)[0], held.get(id));
                } else if (held.containsKey(id) && method.equals("PATCH")) {
                    held.get(id)[0] = between(sent, "\"roles\":[\"", "\"");
                    body = permission(id, held.get(id)[0], held.get(id));
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
    }

    @AfterEach
    void stop() {
        graph.stop(0);
    }

    @Test
    void everyScenarioReplaysWithNoTenantDataInTheRecording(@TempDir Path recordings) throws Exception {
        String base = "http://127.0.0.1:" + graph.getAddress().getPort();
        AdminGrantScenarios.Backend backend = new AdminGrantScenarios.Backend(SITE, GRANTEE, SELF,
                "Contoso Admin", true);

        List<AdminGrantScenarios.Row> rows = AdminGrantScenarios.record(base, () -> "admin-token",
                HttpClient.newHttpClient()::send, backend, recordings);

        assertThat(rows).hasSize(3).allSatisfy(row -> assertThat(row.outcome().status())
                .as(row.scenario().id() + ": " + row.outcome().detail())
                .isEqualTo(AdminGrantScenarios.Status.PASS));
        assertThat(held).as("grants after the run").containsOnlyKeys(permissionId(GRANTEE),
                permissionId(OTHER));
        assertThat(AdminGrantScenarios.violations(recordings)).isEmpty();
        try (Stream<Path> tree = Files.walk(recordings)) {
            for (Path file : tree.filter(Files::isRegularFile).toList()) {
                assertThat(Files.readString(file, StandardCharsets.UTF_8)).as(file.toString())
                        .doesNotContain("contoso", "Contoso", TENANT, GRANTEE, OTHER, SELF,
                                permissionId(GRANTEE), permissionId(OTHER), permissionId(SELF));
            }
        }

        graph.stop(0);
        List<AdminGrantScenarios.Replayed> replayed = AdminGrantScenarios.replay(recordings);
        assertThat(replayed).hasSize(3).allSatisfy(scenario -> {
            assertThat(scenario.recorded()).as(scenario.scenario().id() + " was recorded").isTrue();
            assertThat(scenario.problems()).as("requests that differ from the recording").isEmpty();
            assertThat(scenario.outcome().status()).as(scenario.outcome().detail())
                    .isEqualTo(AdminGrantScenarios.Status.PASS);
        });
    }
}
