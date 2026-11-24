package ai.pipestream.microsoft;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

import ai.pipestream.microsoft.v1.ListSiteChangesRequest;
import ai.pipestream.microsoft.v1.ListSiteChangesResponse;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SiteChange;
import ai.pipestream.microsoft.v1.SiteChangeObject;
import ai.pipestream.microsoft.v1.SiteChangeScope;
import ai.pipestream.microsoft.v1.SiteChangeType;
import com.fasterxml.jackson.databind.JsonNode;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import java.net.URI;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeSet;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

/** ListSiteChanges against the in-process fake: what it asks SharePoint and what it streams. */
class SiteChangesRpcTest {

    private static final String SITE = "5a58bb09-1fba-41c1-8125-69da264370a0";
    private static final String WEB_ID = "712b8fd7-e89d-4f33-9fbb-a6c2b0a9c5e1";
    private static final String LIST_ID = "0f8fad5b-d9cb-469f-a165-70867728950e";
    private static final String WEB = "/sites/Team/_api";
    private static final String FIRST = "1;2;" + WEB_ID + ";638000000000000000;101";
    private static final String SECOND = "1;2;" + WEB_ID + ";638000000000000001;102";
    private static final String THIRD = "1;2;" + WEB_ID + ";638000000000000002;103";

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private String host;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        host = URI.create(fake.baseUrl()).getHost();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl()).build();
        MicrosoftGrpcService service = new MicrosoftGrpcService(config,
                new GraphFiles(new GraphClient(fake.baseUrl(), () -> "graph-token")), 0);
        service.sharePointTokens((session, resource) -> () -> "sharepoint-token");
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
        fake.stub("/sites/s1", "{\"id\":\"s1\",\"webUrl\":\"" + fake.baseUrl() + "/sites/Team\"}");
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        fake.close();
    }

    private List<ListSiteChangesResponse> read(ListSiteChangesRequest.Builder request) {
        List<ListSiteChangesResponse> out = new ArrayList<>();
        stub.listSiteChanges(request.setSiteId("s1").build()).forEachRemaining(out::add);
        return out;
    }

    private static String item(String token, int type) {
        return "{\"odata.type\":\"SP.ChangeItem\",\"ChangeToken\":{\"StringValue\":\"" + token
                + "\"},\"ChangeType\":" + type + ",\"SiteId\":\"" + SITE.toUpperCase(java.util.Locale.ROOT)
                + "\",\"Time\":\"2026-11-24T17:00:00Z\",\"ItemId\":4,\"ListId\":\"" + LIST_ID
                + "\",\"UniqueId\":\"7c9e6679-7425-40de-944b-e07fc1f90ae7\",\"WebId\":\"" + WEB_ID
                + "\",\"ContentTypeId\":{\"StringValue\":\"0x0100AB\"},\"Editor\":\"Ada\"}";
    }

    /** Answers the log in two pages and then nothing, by the token each request starts from. */
    private void twoPages() {
        fake.respond(WEB + "/web/getchanges", request -> {
            String body = request.body();
            if (body.contains(THIRD)) {
                return FakeGraphServer.Stub.json("{\"value\":[]}");
            }
            if (body.contains(SECOND)) {
                return FakeGraphServer.Stub.json("{\"value\":[" + item(THIRD, 18) + "]}");
            }
            return FakeGraphServer.Stub.json("{\"value\":[" + item(FIRST, 17) + ","
                    + "{\"odata.type\":\"SP.ChangeGroup\",\"ChangeToken\":{\"StringValue\":\"" + SECOND
                    + "\"},\"ChangeType\":13,\"SiteId\":\"" + SITE
                    + "\",\"Time\":\"2026-11-24T17:00:01Z\",\"GroupId\":5}]}");
        });
    }

    private JsonNode query(int request) throws Exception {
        return Recordings.JSON.readTree(fake.requestsTo(WEB + "/web/getchanges").get(request).body())
                .path("query");
    }

    @Test
    void theWholeLogIsReadPageByPageAndEndsOnTheLastToken() throws Exception {
        twoPages();

        List<ListSiteChangesResponse> stream = read(ListSiteChangesRequest.newBuilder());

        assertThat(stream).extracting(ListSiteChangesResponse::getEventCase).containsExactly(
                ListSiteChangesResponse.EventCase.CHANGE, ListSiteChangesResponse.EventCase.CHANGE,
                ListSiteChangesResponse.EventCase.CHANGE, ListSiteChangesResponse.EventCase.CHANGE_TOKEN);
        assertThat(stream.get(3).getChangeToken()).isEqualTo(THIRD);
        SiteChange broke = stream.get(0).getChange();
        assertThat(broke.getObject()).isEqualTo(SiteChangeObject.SITE_CHANGE_OBJECT_ITEM);
        assertThat(broke.getType()).isEqualTo(SiteChangeType.SITE_CHANGE_TYPE_SCOPE_ADD);
        assertThat(broke.getTime().getSeconds()).isEqualTo(1795539600L);
        assertThat(broke.getChangeToken()).isEqualTo(FIRST);
        assertThat(broke.getSiteCollectionId()).isEqualTo(SITE);
        assertThat(broke.getWebId()).isEqualTo(WEB_ID);
        assertThat(broke.getSiteId()).isEqualTo(host + "," + SITE + "," + WEB_ID);
        assertThat(broke.getListId()).isEqualTo(LIST_ID);
        assertThat(broke.getItemId()).isEqualTo("4");
        assertThat(broke.getUniqueId()).isEqualTo("7c9e6679-7425-40de-944b-e07fc1f90ae7");
        assertThat(broke.getContentTypeId()).isEqualTo("0x0100AB");
        assertThat(broke.getSecurable().getListItem().getSiteId()).isEqualTo(broke.getSiteId());
        assertThat(broke.getSecurable().getListItem().getListId()).isEqualTo(LIST_ID);
        assertThat(broke.getSecurable().getListItem().getItemId()).isEqualTo("4");
        SiteChange member = stream.get(1).getChange();
        assertThat(member.getObject()).isEqualTo(SiteChangeObject.SITE_CHANGE_OBJECT_GROUP);
        assertThat(member.getType()).isEqualTo(SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_ADD);
        assertThat(member.getGroupId()).isEqualTo(5);
        assertThat(member.getSiteId()).isEmpty();
        assertThat(member.hasSecurable()).isFalse();
        assertThat(stream.get(2).getChange().getType())
                .isEqualTo(SiteChangeType.SITE_CHANGE_TYPE_SCOPE_DELETE);

        assertThat(fake.requestsTo(WEB + "/web/getchanges")).hasSize(3)
                .allMatch(request -> request.authorization().equals("Bearer sharepoint-token"));
        JsonNode first = query(0);
        assertThat(new TreeSet<>(first.properties().stream().map(java.util.Map.Entry::getKey).toList()))
                .containsExactly("Add", "Alert", "ContentType", "DeleteObject", "Field", "File", "Folder",
                        "Group", "GroupMembershipAdd", "GroupMembershipDelete", "Item", "List", "Move",
                        "Navigation", "Rename", "Restore", "RoleAssignmentAdd", "RoleAssignmentDelete",
                        "RoleDefinitionAdd", "RoleDefinitionDelete", "RoleDefinitionUpdate",
                        "SecurityPolicy", "Site", "SystemUpdate", "Update", "User", "View", "Web");
        assertThat(first.properties()).allMatch(flag -> flag.getValue().asBoolean());
        assertThat(query(1).path("ChangeTokenStart").path("StringValue").asText()).isEqualTo(SECOND);
        assertThat(query(2).path("ChangeTokenStart").path("StringValue").asText()).isEqualTo(THIRD);
    }

    @Test
    void aFilterAsksSharePointForItsFlagsAndDropsWhatElseComes() throws Exception {
        twoPages();

        List<ListSiteChangesResponse> stream = read(ListSiteChangesRequest.newBuilder()
                .setChangeToken("1;2;" + WEB_ID + ";637000000000000000;1")
                .addObjects(SiteChangeObject.SITE_CHANGE_OBJECT_GROUP)
                .addTypes(SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_ADD)
                .addTypes(SiteChangeType.SITE_CHANGE_TYPE_MOVE_INTO));

        assertThat(stream).hasSize(2);
        assertThat(stream.get(0).getChange().getGroupId()).isEqualTo(5);
        assertThat(stream.get(1).getChangeToken()).isEqualTo(THIRD);
        JsonNode first = query(0);
        assertThat(new TreeSet<>(first.properties().stream().map(java.util.Map.Entry::getKey).toList()))
                .containsExactly("ChangeTokenStart", "Group", "GroupMembershipAdd", "Move");
    }

    @Test
    void aTypeWithoutAFlagOfItsOwnAsksForEveryType() throws Exception {
        twoPages();

        List<ListSiteChangesResponse> stream = read(ListSiteChangesRequest.newBuilder()
                .addTypes(SiteChangeType.SITE_CHANGE_TYPE_SCOPE_DELETE));

        assertThat(stream).hasSize(2);
        assertThat(stream.get(0).getChange().getChangeToken()).isEqualTo(THIRD);
        assertThat(query(0).has("RoleAssignmentAdd")).isTrue();
        assertThat(query(0).has("Update")).isTrue();
    }

    @Test
    void aLogWithNothingNewEndsOnTheTokenItWasReadFrom() {
        fake.stub(WEB + "/site/getchanges", "{\"value\":[]}");

        List<ListSiteChangesResponse> stream = read(ListSiteChangesRequest.newBuilder()
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION).setChangeToken(FIRST));

        assertThat(stream).hasSize(1);
        assertThat(stream.get(0).getChangeToken()).isEqualTo(FIRST);
    }

    @Test
    void latestAnswersOnlyTheEndOfTheLogOfEachScope() {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(WEB + "/web", "{\"CurrentChangeToken\":{\"StringValue\":\"web-end\"}}");
        fake.stub(WEB + "/site", "{\"CurrentChangeToken\":{\"StringValue\":\"collection-end\"}}");
        fake.stub(list, "{\"CurrentChangeToken\":{\"StringValue\":\"list-end\"}}");

        assertThat(read(ListSiteChangesRequest.newBuilder().setChangeToken("latest")))
                .extracting(ListSiteChangesResponse::getChangeToken).containsExactly("web-end");
        assertThat(read(ListSiteChangesRequest.newBuilder().setChangeToken("latest")
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION)))
                .extracting(ListSiteChangesResponse::getChangeToken).containsExactly("collection-end");
        assertThat(read(ListSiteChangesRequest.newBuilder().setChangeToken("latest")
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_LIST).setListId(LIST_ID)))
                .extracting(ListSiteChangesResponse::getChangeToken).containsExactly("list-end");
        assertThat(fake.only(list).query()).isEqualTo("$select=CurrentChangeToken");
        assertThat(fake.requests()).noneMatch(request -> request.path().endsWith("/getchanges"));
    }

    @Test
    void anEmptyLogReadFromItsStartEndsOnTheCurrentToken() {
        String list = WEB + "/web/lists(guid'" + LIST_ID + "')";
        fake.stub(list + "/getchanges", "{\"value\":[]}");
        fake.stub(list, "{\"CurrentChangeToken\":{\"StringValue\":\"list-end\"}}");

        assertThat(read(ListSiteChangesRequest.newBuilder()
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_LIST).setListId(LIST_ID)))
                .extracting(ListSiteChangesResponse::getChangeToken).containsExactly("list-end");
    }

    @Test
    void aChangeOfAKindThisContractDoesNotKnowKeepsWhatSharePointCalledIt() {
        fake.respond(WEB + "/web/getchanges", request -> FakeGraphServer.Stub.json(
                request.body().contains(FIRST) ? "{\"value\":[]}"
                        : "{\"value\":[{\"odata.type\":\"SP.ChangeSomething\",\"ChangeToken\":"
                                + "{\"StringValue\":\"" + FIRST + "\"},\"ChangeType\":99,\"SiteId\":\""
                                + SITE + "\",\"Time\":\"2026-11-24T17:00:00Z\"}]}"));

        SiteChange change = read(ListSiteChangesRequest.newBuilder()).get(0).getChange();

        assertThat(change.getObject()).isEqualTo(SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED);
        assertThat(change.getType()).isEqualTo(SiteChangeType.SITE_CHANGE_TYPE_UNSPECIFIED);
        assertThat(change.getMetadataMap()).containsEntry("odata.type", "SP.ChangeSomething")
                .containsEntry("ChangeType", "99");
    }

    @Test
    void aScopeAndItsListIdMustAgreeAndFilterValuesMustBeSet() {
        assertInvalid(ListSiteChangesRequest.newBuilder().setScope(SiteChangeScope.SITE_CHANGE_SCOPE_LIST));
        assertInvalid(ListSiteChangesRequest.newBuilder().setListId(LIST_ID));
        assertInvalid(ListSiteChangesRequest.newBuilder()
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_LIST).setListId("not-a-guid"));
        assertInvalid(ListSiteChangesRequest.newBuilder()
                .addObjects(SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED));
        assertInvalid(ListSiteChangesRequest.newBuilder()
                .addTypes(SiteChangeType.SITE_CHANGE_TYPE_UNSPECIFIED));
        assertThat(fake.requests()).noneMatch(request -> request.path().contains("getchanges"));
    }

    private void assertInvalid(ListSiteChangesRequest.Builder request) {
        assertThatThrownBy(() -> read(request)).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.INVALID_ARGUMENT));
    }
}
