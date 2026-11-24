package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.AddSiteGroupMemberRequest;
import ai.pipestream.microsoft.v1.AddSiteRoleAssignmentRequest;
import ai.pipestream.microsoft.v1.CreateSiteGroupRequest;
import ai.pipestream.microsoft.v1.DeleteSiteGroupRequest;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListSiteRoleDefinitionsRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.RemoveSiteGroupMemberRequest;
import ai.pipestream.microsoft.v1.RemoveSiteRoleAssignmentRequest;
import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SitePrincipalType;
import ai.pipestream.microsoft.v1.SiteRoleDefinition;
import ai.pipestream.microsoft.v1.SiteSecurable;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** The rpcs that change SharePoint groups, their members and role assignments, against a fake site. */
class SiteSecurityWriteRpcTest {

    private static final String LIST_ID = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String WEB = "/sites/Team/_api";
    private static final String EMPTY = "{\"odata.null\":true}";

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl()).build();
        MicrosoftGrpcService service = new MicrosoftGrpcService(config,
                new GraphFiles(new GraphClient(fake.baseUrl(), () -> "graph-token")), 0);
        service.sharePointTokens((session, resource) -> () -> "sharepoint-token");
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

    @Test
    void aGroupIsCreatedWithItsTitleAndDescriptionAndDeletedById() {
        fake.stub("POST", WEB + "/web/sitegroups", FakeGraphServer.json(201, """
                {"Id":31,"LoginName":"Reviewers","Title":"Reviewers","PrincipalType":8,
                 "Description":"may review","OwnerTitle":"Team Owners"}""", Map.of()));
        fake.stub("POST", WEB + "/web/sitegroups/removebyid(31)", FakeGraphServer.json(200, EMPTY, Map.of()));

        SiteGroup group = stub.createSiteGroup(CreateSiteGroupRequest.newBuilder().setSiteId("s1")
                .setTitle(" Reviewers ").setDescription("may review").build()).getGroup();
        stub.deleteSiteGroup(DeleteSiteGroupRequest.newBuilder().setSiteId("s1").setGroupId(31).build());

        assertThat(group).isEqualTo(SiteGroup.newBuilder().setId(31).setTitle("Reviewers")
                .setLoginName("Reviewers").setDescription("may review").setOwnerTitle("Team Owners").build());
        FakeGraphServer.RecordedRequest created = fake.only(WEB + "/web/sitegroups");
        assertThat(created.authorization()).isEqualTo("Bearer sharepoint-token");
        assertThat(created.body()).isEqualTo("{\"Title\":\"Reviewers\",\"Description\":\"may review\"}");
        assertThat(fake.requestsTo(WEB + "/web/sitegroups/removebyid(31)")).hasSize(1);
        assertInvalid(() -> stub.createSiteGroup(CreateSiteGroupRequest.newBuilder().setSiteId("s1").build()));
        assertInvalid(() -> stub.deleteSiteGroup(DeleteSiteGroupRequest.newBuilder().setSiteId("s1").build()));
    }

    @Test
    void aMemberIsAddedByLoginNameOrByUserPrincipalNameAndRemovedById() {
        String users = WEB + "/web/sitegroups/getbyid(31)/users";
        fake.stub("POST", users, FakeGraphServer.json(201, """
                {"Id":12,"LoginName":"i:0#.f|membership|ada@contoso.com","Title":"Ada",
                 "PrincipalType":1,"Email":"ada@contoso.com","UserPrincipalName":"ada@contoso.com"}""",
                Map.of()));
        fake.stub("POST", users + "/removebyid(12)", FakeGraphServer.json(200, EMPTY, Map.of()));

        SitePrincipal member = stub.addSiteGroupMember(AddSiteGroupMemberRequest.newBuilder().setSiteId("s1")
                .setGroupId(31).setUserPrincipalName("ada@contoso.com").build()).getMember();
        stub.addSiteGroupMember(AddSiteGroupMemberRequest.newBuilder().setSiteId("s1").setGroupId(31)
                .setLoginName("c:0t.c|tenant|0f8fad5b-d9cb-469f-a165-70867728950e").build());
        stub.removeSiteGroupMember(RemoveSiteGroupMemberRequest.newBuilder().setSiteId("s1")
                .setGroupId(31).setMemberId(12).build());

        assertThat(member).isEqualTo(SitePrincipal.newBuilder().setId(12)
                .setPrincipalType(SitePrincipalType.SITE_PRINCIPAL_TYPE_USER)
                .setLoginName("i:0#.f|membership|ada@contoso.com").setTitle("Ada")
                .setEmail("ada@contoso.com").setUserPrincipalName("ada@contoso.com").build());
        assertThat(fake.requestsTo(users)).extracting(FakeGraphServer.RecordedRequest::body).containsExactly(
                "{\"LoginName\":\"i:0#.f|membership|ada@contoso.com\"}",
                "{\"LoginName\":\"c:0t.c|tenant|0f8fad5b-d9cb-469f-a165-70867728950e\"}");
        assertThat(fake.requestsTo(users + "/removebyid(12)")).hasSize(1);
        assertInvalid(() -> stub.addSiteGroupMember(AddSiteGroupMemberRequest.newBuilder().setSiteId("s1")
                .setGroupId(31).build()));
        assertInvalid(() -> stub.addSiteGroupMember(AddSiteGroupMemberRequest.newBuilder().setSiteId("s1")
                .setLoginName("x").build()));
        assertInvalid(() -> stub.removeSiteGroupMember(RemoveSiteGroupMemberRequest.newBuilder()
                .setSiteId("s1").setGroupId(31).build()));
    }

    @Test
    void thePermissionLevelsOfASiteAreListed() {
        fake.stub(WEB + "/web/roledefinitions", """
                {"value":[{"Id":1073741826,"Name":"Read","Description":"Can view","RoleTypeKind":2,
                  "Hidden":false,"BasePermissions":{"High":"176","Low":"138612833"}}]}""");

        List<SiteRoleDefinition> roles = new ArrayList<>();
        stub.listSiteRoleDefinitions(ListSiteRoleDefinitionsRequest.newBuilder().setSiteId("s1").build())
                .forEachRemaining(response -> roles.add(response.getRole()));

        assertThat(roles).containsExactly(SiteRoleDefinition.newBuilder().setId(1073741826L).setName("Read")
                .setDescription("Can view").setRoleTypeKind(2).setBasePermissionsHigh(176)
                .setBasePermissionsLow(138612833).build());
    }

    @Test
    void aPermissionLevelIsGrantedAndTakenOnASecurable() {
        String assignments = WEB + "/web/lists(guid'" + LIST_ID + "')/items(42)/roleassignments";
        String ids = "(principalid=31,roledefid=1073741826)";
        fake.stub("POST", assignments + "/addroleassignment" + ids, FakeGraphServer.json(200, EMPTY, Map.of()));
        fake.stub("POST", assignments + "/removeroleassignment" + ids,
                FakeGraphServer.json(200, EMPTY, Map.of()));
        SiteSecurable item = SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder()
                .setSiteId("s1").setListId(LIST_ID).setItemId("42")).build();

        stub.addSiteRoleAssignment(AddSiteRoleAssignmentRequest.newBuilder().setSecurable(item)
                .setPrincipalId(31).setRoleDefinitionId(1073741826L).build());
        stub.removeSiteRoleAssignment(RemoveSiteRoleAssignmentRequest.newBuilder().setSecurable(item)
                .setPrincipalId(31).setRoleDefinitionId(1073741826L).build());

        assertThat(fake.requestsTo(assignments + "/addroleassignment" + ids)).hasSize(1);
        assertThat(fake.requestsTo(assignments + "/removeroleassignment" + ids)).hasSize(1);
        assertInvalid(() -> stub.addSiteRoleAssignment(AddSiteRoleAssignmentRequest.newBuilder()
                .setSecurable(item).setRoleDefinitionId(1073741826L).build()));
        assertInvalid(() -> stub.removeSiteRoleAssignment(RemoveSiteRoleAssignmentRequest.newBuilder()
                .setSecurable(item).setPrincipalId(31).build()));
        assertInvalid(() -> stub.addSiteRoleAssignment(AddSiteRoleAssignmentRequest.newBuilder()
                .setPrincipalId(31).setRoleDefinitionId(1073741826L).build()));
    }

    private static void assertInvalid(org.assertj.core.api.ThrowableAssert.ThrowingCallable call) {
        assertThatThrownBy(call).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
