package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;
import static org.assertj.core.api.Assertions.tuple;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItemInclude;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.GetSiteRoleInheritanceRequest;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.SetSiteRoleInheritanceRequest;
import ai.pipestream.microsoft.v1.SiteGroupMembers;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.sync.v1.AclCompleteness;
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
        fake.stub(WEB + "/web", "{\"CurrentChangeToken\":{\"StringValue\":\"t1\"}}");
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
        assertThat(members).extracting(SitePrincipal::getDirectoryGroupOwners)
                .containsExactly(false, false, true, false);
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

    @Test
    void roleInheritanceIsBrokenAndRestoredOnAListItemOnly() {
        String item = WEB + "/web/lists(guid'" + LIST_ID + "')/items(42)";
        fake.stub("POST", item + "/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)",
                FakeGraphServer.json(200, "{}", Map.of()));
        fake.stub("POST", item + "/resetroleinheritance", FakeGraphServer.json(200, "{}", Map.of()));
        SiteSecurable securable = SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder()
                .setSiteId("s1").setListId(LIST_ID).setItemId("42")).build();

        fake.stub(item, "{\"HasUniqueRoleAssignments\":true}");
        assertThat(stub.setSiteRoleInheritance(SetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(securable).setCopyRoleAssignments(true).build())
                .getHasUniqueRoleAssignments()).isTrue();
        fake.stub(item, "{\"HasUniqueRoleAssignments\":false}");
        assertThat(stub.setSiteRoleInheritance(SetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(securable).setInherit(true).build()).getHasUniqueRoleAssignments()).isFalse();

        assertThat(fake.requestsTo(item
                + "/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)")).hasSize(1);
        assertThat(fake.requestsTo(item + "/resetroleinheritance")).hasSize(1);
        assertInvalid(() -> stub.setSiteRoleInheritance(SetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setSiteId("s1")).build()));
    }

    @Test
    void aSecuritySyncEmitsGroupsAclsAndTheOwnerOfEveryItem() {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(WEB + "/web/sitegroups", """
                {"value":[{"Id":5,"LoginName":"Team Owners","Title":"Team Owners","PrincipalType":8}]}""");
        fake.stub(WEB + "/web/sitegroups/getbyid(5)/users", """
                {"value":[{"Id":12,"LoginName":"c:0t.c|tenant|0f8fad5b-d9cb-469f-a165-70867728950e",
                  "Title":"Engineers","PrincipalType":4}]}""");
        fake.stub(WEB + "/web/roleassignments", """
                {"value":[{"PrincipalId":5,"Member":{"Id":5,"LoginName":"Team Owners","Title":"Team Owners",
                  "PrincipalType":8},"RoleDefinitionBindings":[{"Id":1073741829,"Name":"Full Control",
                  "RoleTypeKind":5}]}]}""");
        fake.stub(WEB + "/web/siteusers", """
                {"value":[{"Id":3,"LoginName":"i:0#.f|membership|root@contoso.com","Title":"Root",
                  "PrincipalType":1,"IsSiteAdmin":true}]}""");
        fake.stub("/sites/s1/drives", """
                {"value":[{"id":"drive-1","name":"Docs","driveType":"documentLibrary",
                  "sharePointIds":{"listId":"%s"}}]}""".formatted(LIST_ID));
        fake.stub(list, "{\"HasUniqueRoleAssignments\":false}");
        fake.stub(list + "/items", """
                {"value":[
                 {"Id":1,"HasUniqueRoleAssignments":true,"FileRef":"/sites/Team/Docs/Secret",
                  "FileDirRef":"/sites/Team/Docs"},
                 {"Id":2,"HasUniqueRoleAssignments":false,"FileRef":"/sites/Team/Docs/Secret/in/a.txt",
                  "FileDirRef":"/sites/Team/Docs/Secret/in"},
                 {"Id":3,"HasUniqueRoleAssignments":false,"FileRef":"/sites/Team/Docs/open.txt",
                  "FileDirRef":"/sites/Team/Docs"}]}""");
        fake.stub(list + "/items(1)/roleassignments", """
                {"value":[{"PrincipalId":11,"Member":{"Id":11,"LoginName":"i:0#.f|membership|ada@contoso.com",
                  "Title":"Ada","PrincipalType":1,"Email":"ada@contoso.com"},
                  "RoleDefinitionBindings":[{"Id":1073741826,"Name":"Read","RoleTypeKind":2}]}]}""");
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1",
                MicrosoftFixtures.rootJson("drive-1"), driveItem("folder-1", "1", true),
                driveItem("file-1", "2", false), driveItem("file-2", "3", false)));

        List<MicrosoftEntity> entities = drain(stub.sync(SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build()))
                .stream().filter(SyncResponse::hasChange).map(r -> r.getChange().getEntity()).toList();

        String siteAcl = "site_acl/site/s1";
        String folderAcl = "site_acl/list_item/s1/" + LIST_ID + "/1";
        assertThat(entities).extracting(MicrosoftEntity::getEntityId).containsExactly("s1",
                "site_group/s1/5", "site_admins/s1", siteAcl, folderAcl, "drive-1", "root", "folder-1",
                "file-1", "file-2");
        assertThat(entities.get(2).getSiteAdmins().getCompleteness())
                .isEqualTo(AclCompleteness.ACL_COMPLETENESS_COMPLETE);
        assertThat(entities.get(2).getSiteAdmins().getAdminsList()).extracting(SitePrincipal::getUserPrincipalName)
                .containsExactly("root@contoso.com");
        SiteGroupMembers group = entities.get(1).getSiteGroup();
        assertThat(group.getCompleteness()).isEqualTo(AclCompleteness.ACL_COMPLETENESS_COMPLETE);
        assertThat(group.getMembersList()).extracting(SitePrincipal::getDirectoryObjectId)
                .containsExactly("0f8fad5b-d9cb-469f-a165-70867728950e");
        assertThat(entities.get(3).getSiteAcl().getCompleteness())
                .isEqualTo(AclCompleteness.ACL_COMPLETENESS_COMPLETE);
        assertThat(entities.get(3).getSiteAcl().getAssignments(0).getRoles(0).getName())
                .isEqualTo("Full Control");
        SitePrincipal reader = entities.get(4).getSiteAcl().getAssignments(0).getMember();
        assertThat(reader.getUserPrincipalName()).isEqualTo("ada@contoso.com");
        assertThat(reader.getEmail()).isEqualTo("ada@contoso.com");
        assertThat(entities.get(5).getDrive().getAclOwnerId()).isEqualTo(siteAcl);
        assertThat(entities.subList(6, 10)).extracting(e -> e.getDriveItem().getAclOwnerId())
                .containsExactly(siteAcl, folderAcl, folderAcl, siteAcl);
    }

    @Test
    void aListWhoseSecurityIsRefusedIsMarkedUnavailable() {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(WEB + "/web/sitegroups", "{\"value\":[]}");
        fake.stub(WEB + "/web/roleassignments", "{\"value\":[]}");
        fake.stub("/sites/s1/drives", """
                {"value":[{"id":"drive-1","name":"Docs","driveType":"documentLibrary",
                  "sharePointIds":{"listId":"%s"}}]}""".formatted(LIST_ID));
        fake.stub(list, "{\"HasUniqueRoleAssignments\":false}");
        fake.stub(list + "/items", FakeGraphServer.json(403, """
                {"odata.error":{"code":"-2147024891, System.UnauthorizedAccessException",
                  "message":{"lang":"en-US","value":"Access denied."}}}""", Map.of()));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1", driveItem("file-1", "2", false)));

        List<MicrosoftEntity> entities = drain(stub.sync(SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build()))
                .stream().filter(SyncResponse::hasChange).map(r -> r.getChange().getEntity()).toList();

        String listAcl = "site_acl/list/s1/" + LIST_ID;
        MicrosoftEntity acl = entities.stream().filter(e -> e.getEntityId().equals(listAcl)).findFirst()
                .orElseThrow();
        assertThat(acl.getSiteAcl().getCompleteness()).isEqualTo(AclCompleteness.ACL_COMPLETENESS_UNAVAILABLE);
        assertThat(acl.getSiteAcl().getAssignmentsList()).isEmpty();
        assertThat(entities.getLast().getDriveItem().getAclOwnerId()).isEqualTo(listAcl);
    }

    @Test
    void aResumedSecuritySyncEmitsAgainTheItemsWhoseRoleAssignmentsMoved() {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        String folderAcl = "site_acl/list_item/s1/" + LIST_ID + "/1";
        securedLibrary(true);
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1",
                MicrosoftFixtures.rootJson("drive-1"), driveItem("folder-1", "1", true),
                driveItem("file-1", "2", false), driveItem("file-2", "3", false)));
        SyncRequest request = SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build();
        String cursor = drain(stub.sync(request)).getLast().getResumeCursor();
        assertThat(fake.requests()).noneMatch(r -> r.path().endsWith("/getchanges"));

        // The folder inherits again; delta reports nothing.
        securedLibrary(false);
        fake.respond(WEB + "/web/getchanges", r -> FakeGraphServer.Stub.json(r.body().contains("\"t2\"")
                ? "{\"value\":[]}"
                : "{\"value\":[" + scopeChange("SP.ChangeItem", 18, "t2", ",\"ItemId\":1") + "]}"));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d2"));
        String items = "/sites/s1/lists/" + LIST_ID + "/items/";
        fake.stub(items + "1/driveItem", driveItem("folder-1", "1", true));
        fake.stub(items + "2/driveItem", driveItem("file-1", "2", false));

        List<SyncResponse> resumed = drain(stub.sync(request.toBuilder().setSinceCursor(cursor).build()));

        List<MicrosoftChange> changes = resumed.stream().filter(SyncResponse::hasChange)
                .map(SyncResponse::getChange).toList();
        assertThat(changes).filteredOn(c -> c.getEntity().hasDriveItem())
                .extracting(c -> c.getEntity().getEntityId(), c -> c.getEntity().getDriveItem().getAclOwnerId())
                .containsExactly(tuple("folder-1", "site_acl/site/s1"), tuple("file-1", "site_acl/site/s1"));
        assertThat(changes).filteredOn(c -> c.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE)
                .extracting(c -> c.getEntity().getEntityId()).containsExactly(folderAcl);
        assertThat(fake.requestsTo(WEB + "/web/getchanges").getFirst().body()).contains("\"t1\"");
        assertThat(fake.requestsTo(items + "3/driveItem")).isEmpty();
        assertThat(fake.requestsTo(list + "/items(1)/roleassignments")).hasSize(1);

        // The next pass starts from the end of the log and finds nothing to read again.
        String next = resumed.getLast().getResumeCursor();
        fake.stub(items + "1/driveItem", FakeGraphServer.error(500, "generalException", Map.of()));
        assertThat(drain(stub.sync(request.toBuilder().setSinceCursor(next).build())))
                .filteredOn(SyncResponse::hasChange).noneMatch(r -> r.getChange().getEntity().hasDriveItem());
    }

    @Test
    void aListThatGetsItsOwnRoleAssignmentsIsSweptInFullAgain() {
        securedLibrary(false);
        String page = MicrosoftFixtures.deltaJson(fake.baseUrl() + "/drives/drive-1/root/delta?token=d1",
                MicrosoftFixtures.rootJson("drive-1"), driveItem("file-2", "3", false));
        fake.stub("/drives/drive-1/root/delta", page);
        SyncRequest request = SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build();
        String cursor = drain(stub.sync(request)).getLast().getResumeCursor();

        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(list, "{\"HasUniqueRoleAssignments\":true}");
        fake.stub(list + "/roleassignments", "{\"value\":[]}");
        fake.respond(WEB + "/web/getchanges", r -> FakeGraphServer.Stub.json(r.body().contains("\"t2\"")
                ? "{\"value\":[]}" : "{\"value\":[" + scopeChange("SP.ChangeList", 17, "t2", "") + "]}"));

        List<SyncResponse> resumed = drain(stub.sync(request.toBuilder().setSinceCursor(cursor).build()));

        String listAcl = "site_acl/list/s1/" + LIST_ID;
        assertThat(resumed).filteredOn(SyncResponse::hasChange).map(r -> r.getChange().getEntity())
                .filteredOn(MicrosoftEntity::hasDriveItem)
                .extracting(MicrosoftEntity::getEntityId, e -> e.getDriveItem().getAclOwnerId())
                .containsExactly(tuple("root", listAcl), tuple("file-2", listAcl));
        assertThat(fake.requestsTo("/drives/drive-1/root/delta").getLast().query()).doesNotContain("token=d1");
    }

    @Test
    void aFolderWithMoreItemsBelowItThanAreReadOneByOneSweepsItsListInFullAgain() {
        securedLibrary(false);
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1", driveItem("folder-1", "1", true)));
        SyncRequest request = SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build();
        String cursor = drain(stub.sync(request)).getLast().getResumeCursor();

        StringBuilder items = new StringBuilder("""
                {"Id":1,"HasUniqueRoleAssignments":true,"FileRef":"/sites/Team/Docs/Secret",
                 "FileDirRef":"/sites/Team/Docs"}""");
        for (int id = 2; id <= MicrosoftCrawler.MOVED_READ_LIMIT + 1; id++) {
            items.append("""
                    ,{"Id":%d,"HasUniqueRoleAssignments":false,"FileRef":"/sites/Team/Docs/Secret/%d.txt",
                     "FileDirRef":"/sites/Team/Docs/Secret"}""".formatted(id, id));
        }
        fake.stub(WEB + "/web/lists(guid'" + LIST_ID + "')/items", "{\"value\":[" + items + "]}");
        fake.respond(WEB + "/web/getchanges", r -> FakeGraphServer.Stub.json(r.body().contains("\"t2\"")
                ? "{\"value\":[]}"
                : "{\"value\":[" + scopeChange("SP.ChangeItem", 17, "t2", ",\"ItemId\":1") + "]}"));

        List<SyncResponse> resumed = drain(stub.sync(request.toBuilder().setSinceCursor(cursor).build()));

        assertThat(resumed).filteredOn(SyncResponse::hasChange).map(r -> r.getChange().getEntity())
                .filteredOn(MicrosoftEntity::hasDriveItem)
                .extracting(MicrosoftEntity::getEntityId, e -> e.getDriveItem().getAclOwnerId())
                .containsExactly(tuple("folder-1", "site_acl/list_item/s1/" + LIST_ID + "/1"));
        assertThat(fake.requestsTo("/drives/drive-1/root/delta").getLast().query()).doesNotContain("token=d1");
        assertThat(fake.requests()).noneMatch(r -> r.path().endsWith("/driveItem"));
    }

    @Test
    void aChangeTokenSharePointRefusesSweepsTheSiteInFullAndStartsAtTheEndOfTheLog() {
        securedLibrary(false);
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1", driveItem("file-2", "3", false)));
        SyncRequest request = SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build();
        String cursor = drain(stub.sync(request)).getLast().getResumeCursor();

        fake.stub(WEB + "/web", "{\"CurrentChangeToken\":{\"StringValue\":\"t9\"}}");
        fake.stub("POST", WEB + "/web/getchanges", FakeGraphServer.json(400, """
                {"odata.error":{"code":"-2146232832, Microsoft.SharePoint.SPException",
                  "message":{"lang":"en-US","value":"The changeToken refers to a time before the start of the current change log."}}}""",
                Map.of()));

        List<SyncResponse> resumed = drain(stub.sync(request.toBuilder().setSinceCursor(cursor).build()));

        assertThat(resumed).filteredOn(SyncResponse::hasChange).map(r -> r.getChange().getEntity())
                .filteredOn(MicrosoftEntity::hasDriveItem).extracting(MicrosoftEntity::getEntityId)
                .containsExactly("file-2");
        assertThat(fake.requestsTo("/drives/drive-1/root/delta").getLast().query()).doesNotContain("token=d1");
        assertThat(SyncCursors.decode(resumed.getLast().getResumeCursor(), fake.baseUrl())
                .getSecurityChangeTokensMap()).containsExactly(Map.entry("s1", "t9"));
    }

    @Test
    void aChangeLogTheGrantCannotReadLeavesThePassAsItWas() {
        securedLibrary(false);
        fake.stub(WEB + "/web", FakeGraphServer.json(403, """
                {"odata.error":{"code":"-2147024891, System.UnauthorizedAccessException",
                  "message":{"lang":"en-US","value":"Access denied."}}}""", Map.of()));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/drive-1/root/delta?token=d1", driveItem("file-2", "3", false)));

        List<SyncResponse> pass = drain(stub.sync(SyncRequest.newBuilder().addSiteIds("s1")
                .setInclude(DriveItemInclude.getDefaultInstance()).setIncludeSecurity(true).build()));

        assertThat(pass).filteredOn(SyncResponse::hasChange).map(r -> r.getChange().getEntity())
                .filteredOn(MicrosoftEntity::hasDriveItem).extracting(MicrosoftEntity::getEntityId)
                .containsExactly("file-2");
        assertThat(SyncCursors.decode(pass.getLast().getResumeCursor(), fake.baseUrl())
                .getSecurityChangeTokensMap()).isEmpty();
    }

    /** A site with one library of a folder (item 1), a file in it (2) and a file beside it (3). */
    private void securedLibrary(boolean folderHasItsOwn) {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(WEB + "/web/sitegroups", "{\"value\":[]}");
        fake.stub(WEB + "/web/roleassignments", "{\"value\":[]}");
        fake.stub(WEB + "/web/siteusers", "{\"value\":[]}");
        fake.stub("/sites/s1/drives", """
                {"value":[{"id":"drive-1","name":"Docs","driveType":"documentLibrary",
                  "sharePointIds":{"listId":"%s"}}]}""".formatted(LIST_ID));
        fake.stub(list, "{\"HasUniqueRoleAssignments\":false}");
        fake.stub(list + "/items", """
                {"value":[
                 {"Id":1,"HasUniqueRoleAssignments":%s,"FileRef":"/sites/Team/Docs/Secret",
                  "FileDirRef":"/sites/Team/Docs"},
                 {"Id":2,"HasUniqueRoleAssignments":false,"FileRef":"/sites/Team/Docs/Secret/in/a.txt",
                  "FileDirRef":"/sites/Team/Docs/Secret/in"},
                 {"Id":3,"HasUniqueRoleAssignments":false,"FileRef":"/sites/Team/Docs/Secrets.txt",
                  "FileDirRef":"/sites/Team/Docs"}]}""".formatted(folderHasItsOwn));
        fake.stub(list + "/items(1)/roleassignments", "{\"value\":[]}");
    }

    private static String scopeChange(String type, int changeType, String token, String more) {
        return "{\"odata.type\":\"" + type + "\",\"ChangeToken\":{\"StringValue\":\"" + token
                + "\"},\"ChangeType\":" + changeType
                + ",\"SiteId\":\"7c9e6679-7425-40de-944b-e07fc1f90ae7\""
                + ",\"WebId\":\"3f2504e0-4f89-11d3-9a0c-0305e82c3301\",\"ListId\":\""
                + LIST_ID.toUpperCase(java.util.Locale.ROOT) + "\"" + more + "}";
    }

    private static String driveItem(String id, String listItemId, boolean folder) {
        return """
                {"id":"%s","name":"%s",%s,"parentReference":{"driveId":"drive-1","id":"root"},
                 "sharepointIds":{"listId":"%s","listItemId":"%s"}}"""
                .formatted(id, id, folder ? "\"folder\":{\"childCount\":1}" : "\"file\":{}", LIST_ID, listItemId);
    }

    private static void assertInvalid(org.assertj.core.api.ThrowableAssert.ThrowingCallable call) {
        assertThatThrownBy(call).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
