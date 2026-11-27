package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.ListGroupTransitiveMembersRequest;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListSiteRoleAssignmentsRequest;
import ai.pipestream.microsoft.v1.ListUsersRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SitePrincipalType;
import ai.pipestream.microsoft.v1.SiteRoleAssignment;
import ai.pipestream.microsoft.v1.SiteSecurable;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.net.http.HttpClient;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Records directory and SharePoint site security reads made against fake
 * servers, then replays them through the service without the servers.
 */
class SecurityReadsRecordingTest {

    private static final String SITE = "contoso.sharepoint.com,3f2504e0-4f89-11d3-9a0c-0305e82c3301,"
            + "9b2d5b46-1c2f-4f0e-8a3b-5d7e9c1a2b3c";
    private static final String LIST = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String GROUP = "7c9e6679-7425-40de-944b-e07fc1f90ae7";
    private static final String USER = "1b9d6bcd-bbfd-4b2d-9b5d-ab8dfbbd4bed";

    private final ExecutorService executor = Executors.newCachedThreadPool();
    private final List<Server> servers = new ArrayList<>();
    private final List<ManagedChannel> channels = new ArrayList<>();

    @AfterEach
    void stop() {
        channels.forEach(ManagedChannel::shutdownNow);
        servers.forEach(Server::shutdownNow);
        executor.shutdownNow();
    }

    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub service(String graphBaseUrl,
            GraphTransport transport) throws Exception {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(graphBaseUrl).authority(graphBaseUrl).build();
        MicrosoftGrpcService service = new MicrosoftGrpcService(config, new GraphFiles(
                new GraphClient(graphBaseUrl, () -> "graph-token", AuthMode.AUTH_MODE_APPLICATION,
                        transport)), 0);
        service.sharePointTokens((session, resource) -> () -> "sharepoint-token");
        String name = InProcessServerBuilder.generateName();
        servers.add(InProcessServerBuilder.forName(name).executor(executor).addService(service).build()
                .start());
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        channels.add(channel);
        return MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    /** The reads both the recording and the replaying service answer. */
    private static String reads(MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub, String siteId,
            String listId, String groupId) {
        List<GraphUser> users = new ArrayList<>();
        stub.listUsers(ListUsersRequest.getDefaultInstance())
                .forEachRemaining(response -> users.add(response.getUser()));
        List<String> members = new ArrayList<>();
        stub.listGroupTransitiveMembers(ListGroupTransitiveMembersRequest.newBuilder()
                .setGroupId(groupId).build())
                .forEachRemaining(response -> members.add(response.getMember().getMemberCase().name()));
        List<SiteRoleAssignment> assignments = new ArrayList<>();
        stub.listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId))).build())
                .forEachRemaining(response -> assignments.add(response.getAssignment()));
        SitePrincipal group = assignments.get(0).getMember();
        SitePrincipal entra = assignments.get(1).getMember();
        assertThat(group.getPrincipalType())
                .isEqualTo(SitePrincipalType.SITE_PRINCIPAL_TYPE_SHAREPOINT_GROUP);
        assertThat(group.getLoginName()).isEqualTo(group.getTitle());
        assertThat(entra.getDirectoryObjectId()).hasSize(36);
        assertThat(entra.getLoginName()).endsWith(entra.getDirectoryObjectId());
        return users.size() + " users, first " + (users.get(0).getUserPrincipalName().isEmpty() ? "un" : "")
                + "named, members " + members + ", " + assignments.size() + " assignments, levels "
                + assignments.stream().flatMap(a -> a.getRolesList().stream()).map(r -> r.getName())
                        .distinct().toList();
    }

    @Test
    void directoryAndSiteSecurityReadsReplayWithNoTenantDataInTheRecording(@TempDir Path directory)
            throws Exception {
        FakeGraphServer graph = FakeGraphServer.start();
        FakeGraphServer sharePoint = FakeGraphServer.start();
        try {
            String web = sharePoint.baseUrl() + "/sites/Team";
            graph.stub("/users", """
                    {"value":[{"id":"%s","displayName":"Ada Lovelace","userPrincipalName":"ada@contoso.com",
                      "mail":"ada@contoso.com","givenName":"Ada","surname":"Lovelace","jobTitle":"Engineer",
                      "department":"Research","officeLocation":"Building 9","userType":"Member",
                      "accountEnabled":true,"preferredLanguage":"en-US"}]}""".formatted(USER));
            graph.stub("/groups/" + GROUP + "/transitiveMembers", """
                    {"value":[{"@odata.type":"#microsoft.graph.user","id":"%s","displayName":"Ada Lovelace",
                      "businessPhones":["+1 555 0100"],"mobilePhone":"+1 555 0101",
                      "userPrincipalName":"ada@contoso.com"},
                     {"@odata.type":"#microsoft.graph.group","id":"%s","displayName":"Research Team",
                      "description":"Everyone in research","mailNickname":"research","mailEnabled":false,
                      "securityEnabled":true,"groupTypes":[],"securityIdentifier":"S-1-12-1-11-22-33-44"}]}"""
                    .formatted(USER, GROUP));
            graph.stub("/sites/" + SITE, "{\"id\":\"" + SITE + "\",\"webUrl\":\"" + web + "\"}");
            sharePoint.stub("/sites/Team/_api/web/lists(guid'" + LIST + "')/roleassignments", """
                    {"value":[
                     {"PrincipalId":5,"Member":{"Id":5,"PrincipalType":8,"LoginName":"Team Owners",
                       "Title":"Team Owners","Description":"Owners of Team","OwnerTitle":"Team Owners"},
                      "RoleDefinitionBindings":[{"Id":1073741829,"Name":"Full Control","RoleTypeKind":5,
                        "BasePermissions":{"High":"2147483647","Low":"4294967295"}}]},
                     {"PrincipalId":9,"Member":{"Id":9,"PrincipalType":4,
                       "LoginName":"c:0t.c|tenant|%s","Title":"Research Team","Email":"research@contoso.com",
                       "UserId":{"NameId":"100320012345abcd","NameIdIssuer":"urn:federation:microsoftonline"}},
                      "RoleDefinitionBindings":[{"Id":1073741826,"Name":"Read","RoleTypeKind":2,
                        "BasePermissions":{"High":"176","Low":"138612833"}}]}]}""".formatted(GROUP));

            String base = graph.baseUrl();
            RecordingScrubber scrubber = new RecordingScrubber(base);
            RecordingTransport recorder = new RecordingTransport(HttpClient.newHttpClient()::send,
                    scrubber, new ConcurrentHashMap<>(), base);
            String live = reads(service(base, recorder), SITE, LIST, GROUP);
            assertThat(live).isEqualTo("1 users, first named, members [USER, GROUP], 2 assignments, "
                    + "levels [Full Control, Read]");
            assertThat(recorder.save(directory)).isEqualTo(4);

            StringBuilder recorded = new StringBuilder();
            for (Path file : Recordings.exchanges(directory)) {
                recorded.append(Files.readString(file));
            }
            assertThat(recorded.toString()).doesNotContain("Ada", "Lovelace", "contoso", "Engineer",
                    "Research", "Building", "555 01", "S-1-12-1-11", "Team", "Owners", "100320012345abcd",
                    USER, GROUP, LIST, sharePoint.baseUrl());
            assertThat(RecordingScrubber.violations(recorded.toString())).isEmpty();

            ReplayTransport replay = new ReplayTransport(Map.of());
            replay.begin(directory);
            assertThat(reads(service(base, replay), scrubber.text(SITE), scrubber.text(LIST),
                    scrubber.text(GROUP))).isEqualTo(live);
            assertThat(replay.finish()).isEmpty();
        } finally {
            graph.close();
            sharePoint.close();
        }
    }
}
