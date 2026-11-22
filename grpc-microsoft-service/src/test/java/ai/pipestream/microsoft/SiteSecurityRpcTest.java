package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.GetSiteRoleInheritanceRequest;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListSiteGroupMembersRequest;
import ai.pipestream.microsoft.v1.ListSiteGroupsRequest;
import ai.pipestream.microsoft.v1.ListSiteRoleAssignmentsRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SitePrincipalType;
import ai.pipestream.microsoft.v1.SiteRoleAssignment;
import ai.pipestream.microsoft.v1.SiteSecurable;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

class SiteSecurityRpcTest {

    private static final String LIST_ID = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String WEB = "/sites/Team/_api";

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private MicrosoftGrpcService service;
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private final List<String> resources = new ArrayList<>();

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl()).build();
        service = new MicrosoftGrpcService(config,
                new GraphFiles(new GraphClient(fake.baseUrl(), () -> "graph-token")), 0);
        service.sharePointTokens((session, resource) -> {
            resources.add(resource);
            return () -> "sharepoint-token";
        });
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
        fake.stub("/sites/s1", "{\"id\":\"s1\",\"webUrl\":\"" + fake.baseUrl() + "/sites/Team\"}");
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    private static <T> List<T> drain(Iterator<T> it) {
        List<T> out = new ArrayList<>();
        it.forEachRemaining(out::add);
        return out;
    }

    @Test
    void siteGroupsAreReadFromTheSitesOwnApiWithItsOwnToken() {
        fake.stub(WEB + "/web/sitegroups", """
                {"value":[{"Id":5,"IsHiddenInUI":false,"LoginName":"Team Owners","Title":"Team Owners",
                  "PrincipalType":8,"Description":"full control","OwnerTitle":"Team Owners",
                  "OnlyAllowMembersViewMembership":true}],
                 "odata.nextLink":"web/sitegroups?$skiptoken=2"}""");
        fake.stub(WEB + "/web/sitegroups?$skiptoken=2", """
                {"value":[{"Id":7,"LoginName":"Team Visitors","Title":"Team Visitors","PrincipalType":8}]}""");

        List<SiteGroup> groups = drain(stub.listSiteGroups(
                ListSiteGroupsRequest.newBuilder().setSiteId("s1").build()))
                .stream().map(r -> r.getGroup()).toList();

        assertThat(groups).containsExactly(
                SiteGroup.newBuilder().setId(5).setTitle("Team Owners").setLoginName("Team Owners")
                        .setDescription("full control").setOwnerTitle("Team Owners")
                        .setOnlyMembersViewMembership(true).build(),
                SiteGroup.newBuilder().setId(7).setTitle("Team Visitors").setLoginName("Team Visitors").build());
        assertThat(fake.only("/sites/s1").authorization()).isEqualTo("Bearer graph-token");
        assertThat(fake.requestsTo(WEB + "/web/sitegroups"))
                .extracting(FakeGraphServer.RecordedRequest::authorization)
                .containsExactly("Bearer sharepoint-token", "Bearer sharepoint-token");
        assertThat(resources).containsExactly(fake.baseUrl());
    }

    @Test
    void groupMembersKeepTheirKindAndTheEntraObjectIdOfAGroup() {
        fake.stub(WEB + "/web/sitegroups/getbyid(5)/users", """
                {"value":[
                 {"Id":11,"LoginName":"i:0#.f|membership|ada@contoso.com","Title":"Ada","PrincipalType":1,
                  "Email":"ada@contoso.com","UserPrincipalName":"ada@contoso.com","IsSiteAdmin":true},
                 {"Id":12,"LoginName":"c:0t.c|tenant|0f8fad5b-d9cb-469f-a165-70867728950e","Title":"Engineers",
                  "PrincipalType":4},
                 {"Id":13,"LoginName":"c:0o.c|federateddirectoryclaimprovider|7c9e6679-7425-40de-944b-e07fc1f90ae7_o",
                  "Title":"Team Owners (M365)","PrincipalType":4},
                 {"Id":14,"LoginName":"c:0-.f|rolemanager|spo-grid-all-users/tenant-1","Title":"Everyone except external users",
                  "PrincipalType":4}]}""");

        List<SitePrincipal> members = drain(stub.listSiteGroupMembers(ListSiteGroupMembersRequest.newBuilder()
                .setSiteId("s1").setGroupId(5).build())).stream().map(r -> r.getMember()).toList();

        assertThat(members).extracting(SitePrincipal::getPrincipalType).containsExactly(
                SitePrincipalType.SITE_PRINCIPAL_TYPE_USER, SitePrincipalType.SITE_PRINCIPAL_TYPE_SECURITY_GROUP,
                SitePrincipalType.SITE_PRINCIPAL_TYPE_SECURITY_GROUP,
                SitePrincipalType.SITE_PRINCIPAL_TYPE_SECURITY_GROUP);
        assertThat(members.get(0).getUserPrincipalName()).isEqualTo("ada@contoso.com");
        assertThat(members.get(0).getSiteAdmin()).isTrue();
        assertThat(members).extracting(SitePrincipal::getDirectoryObjectId).containsExactly("",
                "0f8fad5b-d9cb-469f-a165-70867728950e", "7c9e6679-7425-40de-944b-e07fc1f90ae7", "");
        assertInvalid(() -> drain(stub.listSiteGroupMembers(
                ListSiteGroupMembersRequest.newBuilder().setSiteId("s1").build())));
    }

    @Test
    void roleAssignmentsCarryThePrincipalAndItsPermissionLevels() {
        String body = """
                {"value":[{"PrincipalId":5,
                  "Member":{"Id":5,"LoginName":"Team Owners","Title":"Team Owners","PrincipalType":8},
                  "RoleDefinitionBindings":[{"Id":1073741829,"Name":"Full Control","Description":"all",
                    "RoleTypeKind":5,"Hidden":false,
                    "BasePermissions":{"High":"2147483647","Low":"4294967295"}}]}]}""";
        fake.stub(WEB + "/web/roleassignments", body);
        fake.stub(WEB + "/web/lists(guid'" + LIST_ID + "')/roleassignments", body);
        fake.stub(WEB + "/web/lists(guid'" + LIST_ID + "')/items(42)/roleassignments", body);

        SiteRoleAssignment onSite = drain(stub.listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setSiteId("s1")).build())).get(0).getAssignment();
        assertThat(onSite.getPrincipalId()).isEqualTo(5);
        assertThat(onSite.getMember().getPrincipalType())
                .isEqualTo(SitePrincipalType.SITE_PRINCIPAL_TYPE_SHAREPOINT_GROUP);
        assertThat(onSite.getRolesList()).singleElement().satisfies(role -> {
            assertThat(role.getId()).isEqualTo(1073741829L);
            assertThat(role.getName()).isEqualTo("Full Control");
            assertThat(role.getRoleTypeKind()).isEqualTo(5);
            assertThat(Integer.toUnsignedLong(role.getBasePermissionsHigh())).isEqualTo(2147483647L);
            assertThat(Integer.toUnsignedLong(role.getBasePermissionsLow())).isEqualTo(4294967295L);
        });
        assertThat(fake.only(WEB + "/web/roleassignments").query())
                .isEqualTo("$expand=Member,RoleDefinitionBindings");

        assertThat(drain(stub.listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setList(ListRef.newBuilder().setSiteId("s1")
                        .setListId(LIST_ID))).build()))).hasSize(1);
        assertThat(drain(stub.listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder().setSiteId("s1")
                        .setListId(LIST_ID).setItemId("42"))).build()))).hasSize(1);
        assertInvalid(() -> drain(stub.listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setList(ListRef.newBuilder().setSiteId("s1")
                        .setListId("Documents"))).build())));
        assertInvalid(() -> drain(stub.listSiteRoleAssignments(
                ListSiteRoleAssignmentsRequest.getDefaultInstance())));
    }

    @Test
    void inheritanceIsReadPerSecurable() {
        fake.stub(WEB + "/web", "{\"HasUniqueRoleAssignments\":true}");
        fake.stub(WEB + "/web/lists(guid'" + LIST_ID + "')/items(42)", "{\"HasUniqueRoleAssignments\":false}");

        assertThat(stub.getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setSiteId("s1")).build())
                .getHasUniqueRoleAssignments()).isTrue();
        assertThat(stub.getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder().setSiteId("s1")
                        .setListId(LIST_ID).setItemId("42"))).build()).getHasUniqueRoleAssignments()).isFalse();
        assertThat(fake.only(WEB + "/web").query()).isEqualTo("$select=HasUniqueRoleAssignments");
    }

    @Test
    void aSharePointRefusalKeepsItsCodeAndMessage() {
        fake.stub(WEB + "/web/sitegroups", FakeGraphServer.json(403, """
                {"odata.error":{"code":"-2147024891, System.UnauthorizedAccessException",
                  "message":{"lang":"en-US","value":"Access denied."}}}""", Map.of()));

        assertThatThrownBy(() -> drain(stub.listSiteGroups(
                ListSiteGroupsRequest.newBuilder().setSiteId("s1").build())))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> {
                    assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.PERMISSION_DENIED);
                    assertThat(e.getStatus().getDescription()).contains("Access denied.");
                });
    }

    @Test
    void withoutACertificateTheRpcSaysWhatIsMissing() {
        service.sharePointTokens(SharePointRest::certificateTokens);

        assertThatThrownBy(() -> drain(stub.listSiteGroups(
                ListSiteGroupsRequest.newBuilder().setSiteId("s1").build())))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> {
                    assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.FAILED_PRECONDITION);
                    assertThat(e.getStatus().getDescription()).contains("MICROSOFT_CLIENT_CERT_PEM");
                });
        assertThat(fake.requestsTo(WEB + "/web/sitegroups")).isEmpty();
    }

    private static void assertInvalid(org.assertj.core.api.ThrowableAssert.ThrowingCallable call) {
        assertThatThrownBy(call).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
