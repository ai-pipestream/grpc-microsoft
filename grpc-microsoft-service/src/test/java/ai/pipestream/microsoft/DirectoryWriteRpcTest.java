package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.AddGroupMemberRequest;
import ai.pipestream.microsoft.v1.CreateGroupRequest;
import ai.pipestream.microsoft.v1.DeleteGroupRequest;
import ai.pipestream.microsoft.v1.GetGroupRequest;
import ai.pipestream.microsoft.v1.GraphGroup;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.RemoveGroupMemberRequest;
import ai.pipestream.microsoft.v1.UpdateGroupRequest;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import org.assertj.core.api.ThrowableAssert.ThrowingCallable;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** The rpcs that read one Entra group and change groups and their members, against a fake Graph. */
class DirectoryWriteRpcTest {

    private static final String GROUP = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String MEMBER = "7c9e6679-7425-40de-944b-e07fc1f90ae7";

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
    void aGroupIsCreatedAsASecurityGroupUnlessTheRequestSaysOtherwise() {
        fake.stub("POST", "/groups", FakeGraphServer.json(201, """
                {"id":"%s","displayName":"Reviewers","description":"may review","mailNickname":"reviewers",
                 "mailEnabled":false,"securityEnabled":true,"groupTypes":[],
                 "securityIdentifier":"S-1-12-1-1"}""".formatted(GROUP), Map.of()));

        GraphGroup created = stub.createGroup(CreateGroupRequest.newBuilder().setGroup(GraphGroup.newBuilder()
                .setDisplayName(" Reviewers ").setMailNickname("reviewers").setDescription("may review"))
                .build()).getGroup();
        stub.createGroup(CreateGroupRequest.newBuilder().setGroup(GraphGroup.newBuilder()
                .setDisplayName("Team").setMailNickname("team").setMailEnabled(true).setSecurityEnabled(false)
                .addGroupTypes("Unified").setVisibility("Private")).build());

        assertThat(created).isEqualTo(GraphGroup.newBuilder().setId(GROUP).setDisplayName("Reviewers")
                .setDescription("may review").setMailNickname("reviewers").setMailEnabled(false)
                .setSecurityEnabled(true).setSecurityIdentifier("S-1-12-1-1").build());
        assertThat(fake.requestsTo("/groups")).extracting(FakeGraphServer.RecordedRequest::body)
                .containsExactly(
                        "{\"displayName\":\"Reviewers\",\"mailNickname\":\"reviewers\",\"mailEnabled\":false,"
                                + "\"securityEnabled\":true,\"description\":\"may review\",\"groupTypes\":[]}",
                        "{\"displayName\":\"Team\",\"mailNickname\":\"team\",\"mailEnabled\":true,"
                                + "\"securityEnabled\":false,\"visibility\":\"Private\","
                                + "\"groupTypes\":[\"Unified\"]}");
        assertThat(fake.requestsTo("/groups").getFirst().authorization()).isEqualTo("Bearer graph-token");
    }

    @Test
    void aGroupWithoutANameOrAliasOrWithFieldsGraphSetsIsRefused() {
        assertInvalid(() -> stub.createGroup(CreateGroupRequest.newBuilder()
                .setGroup(GraphGroup.newBuilder().setMailNickname("reviewers")).build()));
        assertInvalid(() -> stub.createGroup(CreateGroupRequest.newBuilder()
                .setGroup(GraphGroup.newBuilder().setDisplayName("Reviewers")).build()));
        assertInvalid(() -> stub.createGroup(CreateGroupRequest.newBuilder().setGroup(GraphGroup.newBuilder()
                .setDisplayName("Reviewers").setMailNickname("reviewers").setId(GROUP)).build()));
        assertThat(fake.requests()).isEmpty();
    }

    @Test
    void aGroupIsReadChangedAndDeletedById() {
        fake.stub("/groups/" + GROUP, """
                {"id":"%s","displayName":"Reviewers","mailNickname":"reviewers","mailEnabled":false,
                 "securityEnabled":true,"groupTypes":[]}""".formatted(GROUP));
        fake.stub("PATCH", "/groups/" + GROUP, FakeGraphServer.empty(204, Map.of()));
        fake.stub("DELETE", "/groups/" + GROUP, FakeGraphServer.empty(204, Map.of()));

        GraphGroup read = stub.getGroup(GetGroupRequest.newBuilder().setGroupId(GROUP).build()).getGroup();
        stub.updateGroup(UpdateGroupRequest.newBuilder().setGroupId(GROUP).setDisplayName("Readers")
                .setDescription("").build());
        stub.deleteGroup(DeleteGroupRequest.newBuilder().setGroupId(GROUP).build());

        assertThat(read.getDisplayName()).isEqualTo("Reviewers");
        assertThat(fake.requestsTo("/groups/" + GROUP))
                .extracting(FakeGraphServer.RecordedRequest::method, FakeGraphServer.RecordedRequest::body)
                .containsExactly(org.assertj.core.api.Assertions.tuple("GET", ""),
                        org.assertj.core.api.Assertions.tuple("PATCH",
                                "{\"displayName\":\"Readers\",\"description\":null}"),
                        org.assertj.core.api.Assertions.tuple("DELETE", ""));
        assertInvalid(() -> stub.updateGroup(UpdateGroupRequest.newBuilder().setGroupId(GROUP).build()));
        assertInvalid(() -> stub.updateGroup(UpdateGroupRequest.newBuilder().setGroupId(GROUP)
                .setDisplayName(" ").build()));
        assertInvalid(() -> stub.getGroup(GetGroupRequest.getDefaultInstance()));
        assertInvalid(() -> stub.deleteGroup(DeleteGroupRequest.getDefaultInstance()));
    }

    @Test
    void aMemberIsAddedByReferenceAndRemovedById() {
        fake.stub("POST", "/groups/" + GROUP + "/members/$ref", FakeGraphServer.empty(204, Map.of()));
        fake.stub("DELETE", "/groups/" + GROUP + "/members/" + MEMBER + "/$ref",
                FakeGraphServer.empty(204, Map.of()));

        stub.addGroupMember(AddGroupMemberRequest.newBuilder().setGroupId(GROUP).setMemberId(MEMBER).build());
        stub.removeGroupMember(RemoveGroupMemberRequest.newBuilder().setGroupId(GROUP).setMemberId(MEMBER)
                .build());

        assertThat(fake.only("/groups/" + GROUP + "/members/$ref").body())
                .isEqualTo("{\"@odata.id\":\"" + fake.baseUrl() + "/directoryObjects/" + MEMBER + "\"}");
        assertThat(fake.only("/groups/" + GROUP + "/members/" + MEMBER + "/$ref").method())
                .isEqualTo("DELETE");
        assertInvalid(() -> stub.addGroupMember(AddGroupMemberRequest.newBuilder().setGroupId(GROUP).build()));
        assertInvalid(() -> stub.removeGroupMember(RemoveGroupMemberRequest.newBuilder()
                .setMemberId(MEMBER).build()));
    }

    @Test
    void anAppWithoutTheWritePermissionIsDenied() {
        fake.stub("POST", "/groups", FakeGraphServer.error(403, "Authorization_RequestDenied", Map.of()));

        assertThatThrownBy(() -> stub.createGroup(CreateGroupRequest.newBuilder().setGroup(GraphGroup.newBuilder()
                .setDisplayName("Reviewers").setMailNickname("reviewers")).build()))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(e.getStatus().getCode())
                        .isEqualTo(Status.Code.PERMISSION_DENIED));
    }

    private static void assertInvalid(ThrowingCallable call) {
        assertThatThrownBy(call).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
