package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.CreateSharingLinkRequest;
import ai.pipestream.microsoft.v1.CreateSharingLinkResponse;
import ai.pipestream.microsoft.v1.DeletePermissionRequest;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.GetPermissionRequest;
import ai.pipestream.microsoft.v1.GrantLinkAccessRequest;
import ai.pipestream.microsoft.v1.InviteRecipientsRequest;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.RevokeLinkAccessRequest;
import ai.pipestream.microsoft.v1.SharingRecipient;
import ai.pipestream.microsoft.v1.UpdatePermissionRequest;
import com.google.protobuf.Timestamp;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.assertj.core.api.ThrowableAssert.ThrowingCallable;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** The rpcs that create sharing links and grant, change and remove permissions, against a fake Graph. */
class PermissionWriteRpcTest {

    private static final String ITEM = "/drives/d1/items/i1";
    private static final String LIST_ITEM = "/sites/s1/lists/l1/items/7";
    private static final String USER = "7c9e6679-7425-40de-944b-e07fc1f90ae7";
    private static final DriveItemRef ITEM_REF = DriveItemRef.newBuilder().setDriveId("d1").setItemId("i1").build();
    private static final PermissionTarget ON_ITEM = PermissionTarget.newBuilder().setDriveItem(ITEM_REF).build();
    private static final PermissionTarget ON_LIST_ITEM = PermissionTarget.newBuilder()
            .setListItem(ListItemRef.newBuilder().setSiteId("s1").setListId("l1").setItemId("7")).build();
    private static final String LINK = """
            {"id":"p1","roles":["read"],"shareId":"u!abc",
             "link":{"type":"view","scope":"organization","webUrl":"https://contoso.example/:t:/s/x"}}""";

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
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    @Test
    void aLinkIsCreatedOnADriveItemAndTheAnswerSaysWhetherItIsNew() {
        fake.stub("POST", ITEM + "/createLink", FakeGraphServer.json(201, LINK, Map.of()));
        fake.stub("POST", LIST_ITEM + "/createLink", FakeGraphServer.json(200, LINK, Map.of()));

        CreateSharingLinkResponse made = stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(ON_ITEM).setType(" view ").setScope("organization").build());
        CreateSharingLinkResponse found = stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(ON_LIST_ITEM).setType("edit").setScope("users")
                .setExpiresAt(Timestamp.newBuilder().setSeconds(1_900_000_000L))
                .setRetainInheritedPermissions(false)
                .addRecipients(SharingRecipient.newBuilder().setObjectId(USER))
                .addRecipients(SharingRecipient.newBuilder().setEmail("ada@contoso.example"))
                .setSendNotification(true).build());

        assertThat(made.getCreated()).isTrue();
        assertThat(found.getCreated()).isFalse();
        Permission permission = made.getPermission();
        assertThat(permission.getId()).isEqualTo("p1");
        assertThat(permission.getLink().getScope()).isEqualTo("organization");
        assertThat(permission.getTarget()).isEqualTo(ON_ITEM);
        assertThat(found.getPermission().getTarget()).isEqualTo(ON_LIST_ITEM);
        assertThat(fake.only(ITEM + "/createLink").body())
                .isEqualTo("{\"type\":\"view\",\"scope\":\"organization\"}");
        assertThat(fake.only(LIST_ITEM + "/createLink").body()).isEqualTo(
                "{\"type\":\"edit\",\"scope\":\"users\",\"expirationDateTime\":\"2030-03-17T17:46:40Z\","
                        + "\"retainInheritedPermissions\":false,\"recipients\":[{\"objectId\":\"" + USER
                        + "\"},{\"email\":\"ada@contoso.example\"}],\"sendNotification\":true}");
    }

    @Test
    void aLinkNeedsATypeAnItemAndRecipientsWhenItIsMailed() {
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(ON_ITEM).build()));
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setType("view").build()));
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder().setType("view")
                .setTarget(PermissionTarget.newBuilder().setSiteId("s1")).build()));
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder().setType("view")
                .setTarget(PermissionTarget.newBuilder().setList(ListRef.newBuilder().setSiteId("s1")
                        .setListId("l1"))).build()));
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(ON_ITEM).setType("view").setSendNotification(true).build()));
        assertInvalid(() -> stub.createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(ON_ITEM).setType("view").addRecipients(SharingRecipient.getDefaultInstance())
                .build()));
        assertThat(fake.requests()).isEmpty();
    }

    @Test
    void anInvitationGrantsRolesWithoutMailUnlessAsked() {
        fake.stub("POST", ITEM + "/invite", FakeGraphServer.json(200, """
                {"value":[{"id":"p2","roles":["write"],
                  "grantedToV2":{"user":{"id":"%s","displayName":"Ada"}}}]}""".formatted(USER), Map.of()));

        List<Permission> granted = stub.inviteRecipients(InviteRecipientsRequest.newBuilder().setItem(ITEM_REF)
                .addRecipients(SharingRecipient.newBuilder().setObjectId(USER)).addRoles("write")
                .setRequireSignIn(true).build()).getPermissionsList();
        stub.inviteRecipients(InviteRecipientsRequest.newBuilder().setItem(ITEM_REF)
                .addRecipients(SharingRecipient.newBuilder().setAlias("ada")).addRoles("read")
                .setSendInvitation(true).setMessage("please read")
                .setRetainInheritedPermissions(true).build());

        assertThat(granted).singleElement().satisfies(permission -> {
            assertThat(permission.getId()).isEqualTo("p2");
            assertThat(permission.getGrantedTo().getUser().getId()).isEqualTo(USER);
            assertThat(permission.getTarget()).isEqualTo(ON_ITEM);
        });
        assertThat(fake.requestsTo(ITEM + "/invite")).extracting(FakeGraphServer.RecordedRequest::body)
                .containsExactly(
                        "{\"recipients\":[{\"objectId\":\"" + USER + "\"}],\"roles\":[\"write\"],"
                                + "\"requireSignIn\":true,\"sendInvitation\":false}",
                        "{\"recipients\":[{\"alias\":\"ada\"}],\"roles\":[\"read\"],\"requireSignIn\":false,"
                                + "\"sendInvitation\":true,\"message\":\"please read\","
                                + "\"retainInheritedPermissions\":true}");
        assertInvalid(() -> stub.inviteRecipients(InviteRecipientsRequest.newBuilder().setItem(ITEM_REF)
                .addRoles("read").build()));
        assertInvalid(() -> stub.inviteRecipients(InviteRecipientsRequest.newBuilder().setItem(ITEM_REF)
                .addRecipients(SharingRecipient.newBuilder().setObjectId(USER)).build()));
        assertInvalid(() -> stub.inviteRecipients(InviteRecipientsRequest.newBuilder()
                .addRecipients(SharingRecipient.newBuilder().setObjectId(USER)).addRoles("read").build()));
    }

    @Test
    void aPermissionIsReadChangedAndRemovedById() {
        fake.stub(ITEM + "/permissions/p1", LINK);
        fake.stub("/sites/s1/permissions/g1", "{\"id\":\"g1\",\"roles\":[\"fullcontrol\"]}");
        fake.stub("PATCH", ITEM + "/permissions/p1", FakeGraphServer.json(200, """
                {"id":"p1","roles":["write"]}""", Map.of()));
        fake.stub("DELETE", ITEM + "/permissions/p1", FakeGraphServer.empty(204, Map.of()));
        fake.stub("DELETE", LIST_ITEM + "/permissions/p9", FakeGraphServer.empty(204, Map.of()));

        Permission read = stub.getPermission(GetPermissionRequest.newBuilder().setTarget(ON_ITEM)
                .setPermissionId("p1").build()).getPermission();
        Permission onSite = stub.getPermission(GetPermissionRequest.newBuilder()
                .setTarget(PermissionTarget.newBuilder().setSiteId("s1")).setPermissionId("g1").build())
                .getPermission();
        Permission changed = stub.updatePermission(UpdatePermissionRequest.newBuilder().setTarget(ON_ITEM)
                .setPermissionId("p1").addRoles("write").setClearExpiresAt(true).build()).getPermission();
        stub.updatePermission(UpdatePermissionRequest.newBuilder().setTarget(ON_ITEM).setPermissionId("p1")
                .setExpiresAt(Timestamp.newBuilder().setSeconds(1_900_000_000L)).setPassword("").build());
        stub.deletePermission(DeletePermissionRequest.newBuilder().setTarget(ON_ITEM).setPermissionId("p1")
                .build());
        stub.deletePermission(DeletePermissionRequest.newBuilder().setTarget(ON_LIST_ITEM)
                .setPermissionId("p9").build());

        assertThat(read.getShareId()).isEqualTo("u!abc");
        assertThat(read.getTarget()).isEqualTo(ON_ITEM);
        assertThat(onSite.getRolesList()).containsExactly("fullcontrol");
        assertThat(changed.getRolesList()).containsExactly("write");
        assertThat(fake.requestsTo(ITEM + "/permissions/p1"))
                .extracting(FakeGraphServer.RecordedRequest::method, FakeGraphServer.RecordedRequest::body)
                .containsExactly(org.assertj.core.api.Assertions.tuple("GET", ""),
                        org.assertj.core.api.Assertions.tuple("PATCH",
                                "{\"roles\":[\"write\"],\"expirationDateTime\":null}"),
                        org.assertj.core.api.Assertions.tuple("PATCH",
                                "{\"expirationDateTime\":\"2030-03-17T17:46:40Z\",\"password\":\"\"}"),
                        org.assertj.core.api.Assertions.tuple("DELETE", ""));
        assertThat(fake.only(LIST_ITEM + "/permissions/p9").method()).isEqualTo("DELETE");
    }

    @Test
    void aChangeOrARemovalNeedsAnItemAnIdAndSomethingToChange() {
        assertInvalid(() -> stub.getPermission(GetPermissionRequest.newBuilder().setTarget(ON_ITEM).build()));
        assertInvalid(() -> stub.getPermission(GetPermissionRequest.newBuilder().setPermissionId("p1").build()));
        assertInvalid(() -> stub.updatePermission(UpdatePermissionRequest.newBuilder().setTarget(ON_ITEM)
                .setPermissionId("p1").build()));
        assertInvalid(() -> stub.updatePermission(UpdatePermissionRequest.newBuilder().setTarget(ON_ITEM)
                .setPermissionId("p1").setClearExpiresAt(true)
                .setExpiresAt(Timestamp.newBuilder().setSeconds(1_900_000_000L)).build()));
        assertInvalid(() -> stub.updatePermission(UpdatePermissionRequest.newBuilder()
                .setTarget(PermissionTarget.newBuilder().setSiteId("s1")).setPermissionId("g1")
                .addRoles("read").build()));
        assertInvalid(() -> stub.deletePermission(DeletePermissionRequest.newBuilder()
                .setTarget(PermissionTarget.newBuilder().setSiteId("s1")).setPermissionId("g1").build()));
        assertInvalid(() -> stub.deletePermission(DeletePermissionRequest.newBuilder().setTarget(ON_ITEM)
                .build()));
        assertThat(fake.requests()).isEmpty();
    }

    @Test
    void peopleAreAddedToAndTakenOffALink() {
        fake.stub("POST", ITEM + "/permissions/p1/grant", FakeGraphServer.json(200, """
                {"value":[{"id":"p1","roles":["read"],"link":{"type":"view","scope":"users"},
                  "grantedToIdentitiesV2":[{"user":{"id":"%s"}}]}]}""".formatted(USER), Map.of()));
        fake.stub("POST", ITEM + "/permissions/p1/revokeGrants", FakeGraphServer.json(200, """
                {"id":"p1","roles":["read"],"link":{"type":"view","scope":"users"},
                 "grantedToIdentitiesV2":[]}""", Map.of()));

        List<Permission> granted = stub.grantLinkAccess(GrantLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .setPermissionId("p1").addRecipients(SharingRecipient.newBuilder().setObjectId(USER))
                .addRoles("read").build()).getPermissionsList();
        Permission after = stub.revokeLinkAccess(RevokeLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .setPermissionId("p1").addGrantees(SharingRecipient.newBuilder().setEmail("ada@contoso.example"))
                .build()).getPermission();

        assertThat(granted).singleElement().satisfies(permission -> {
            assertThat(permission.getGrantedToIdentitiesList()).singleElement()
                    .satisfies(identity -> assertThat(identity.getUser().getId()).isEqualTo(USER));
            assertThat(permission.getTarget()).isEqualTo(ON_ITEM);
        });
        assertThat(after.getGrantedToIdentitiesList()).isEmpty();
        assertThat(after.getTarget()).isEqualTo(ON_ITEM);
        assertThat(fake.only(ITEM + "/permissions/p1/grant").body())
                .isEqualTo("{\"recipients\":[{\"objectId\":\"" + USER + "\"}],\"roles\":[\"read\"]}");
        assertThat(fake.only(ITEM + "/permissions/p1/revokeGrants").body())
                .isEqualTo("{\"grantees\":[{\"email\":\"ada@contoso.example\"}]}");
        assertInvalid(() -> stub.grantLinkAccess(GrantLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .setPermissionId("p1").addRoles("read").build()));
        assertInvalid(() -> stub.grantLinkAccess(GrantLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .setPermissionId("p1").addRecipients(SharingRecipient.newBuilder().setObjectId(USER)).build()));
        assertInvalid(() -> stub.revokeLinkAccess(RevokeLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .setPermissionId("p1").build()));
        assertInvalid(() -> stub.revokeLinkAccess(RevokeLinkAccessRequest.newBuilder().setItem(ITEM_REF)
                .addGrantees(SharingRecipient.newBuilder().setObjectId(USER)).build()));
    }

    @Test
    void aRefusalKeepsItsStatus() {
        fake.stub("POST", ITEM + "/createLink", FakeGraphServer.error(403, "accessDenied", Map.of()));
        fake.stub("DELETE", ITEM + "/permissions/p1", FakeGraphServer.error(404, "itemNotFound", Map.of()));

        assertStatus(Status.Code.PERMISSION_DENIED, () -> stub.createSharingLink(
                CreateSharingLinkRequest.newBuilder().setTarget(ON_ITEM).setType("view").build()));
        assertStatus(Status.Code.NOT_FOUND, () -> stub.deletePermission(DeletePermissionRequest.newBuilder()
                .setTarget(ON_ITEM).setPermissionId("p1").build()));
    }

    private static void assertInvalid(ThrowingCallable call) {
        assertStatus(Status.Code.INVALID_ARGUMENT, call);
    }

    private static void assertStatus(Status.Code code, ThrowingCallable call) {
        assertThatThrownBy(call).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(code));
    }
}
