package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.SiteChange;
import ai.pipestream.microsoft.v1.SiteChangeObject;
import ai.pipestream.microsoft.v1.SiteChangeScope;
import ai.pipestream.microsoft.v1.SiteChangeSnapshot;
import ai.pipestream.microsoft.v1.SiteChangeType;
import ai.pipestream.microsoft.v1.SiteChangeWatch;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.microsoft.v1.WatchChangesRequest;
import ai.pipestream.microsoft.v1.WatchChangesResponse;
import io.grpc.Context;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.time.Clock;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * The SharePoint change log part of a WatchChanges pass against the in-process
 * fake: what it asks for, what it streams and what the resume cursor carries.
 */
class MicrosoftChangeSiteLogTest {

    private static final String WEB = "/sites/Team/_api";
    private static final String SITE = "5a58bb09-1fba-41c1-8125-69da264370a0";
    private static final String WEB_ID = "712b8fd7-e89d-4f33-9fbb-a6c2b0a9c5e1";
    private static final String LIST_ID = "0f8fad5b-d9cb-469f-a165-70867728950e";


    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub stub;
    private final List<String> syncedFrom = new CopyOnWriteArrayList<>();
    private String driveLink;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        driveLink = fake.baseUrl() + "/drives/d1/root/delta?token=d";
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl()).build();
        MicrosoftGrpcService microsoft = new MicrosoftGrpcService(config, new GraphFiles(
                new GraphClient(fake.baseUrl(), () -> "token", AuthMode.AUTH_MODE_APPLICATION)), 0);
        microsoft.sharePointTokens((session, resource) -> () -> "sharepoint-token");
        fake.stub("/sites/s1", "{\"id\":\"s1\",\"webUrl\":\"" + fake.baseUrl() + "/sites/Team\"}");
        // Sync is replaced: each pass ends on a cursor that holds one drive link only.
        MicrosoftChangeGrpcService service = new MicrosoftChangeGrpcService(microsoft,
                (request, observer) -> {
                    syncedFrom.add(request.getSinceCursor());
                    observer.onNext(SyncResponse.newBuilder().setResumeCursor(SyncCursors.encode(
                            SyncCursor.newBuilder().putDriveDeltaLinks("d1", driveLink).build())).build());
                    observer.onCompleted();
                }, new NotificationBus(), false, "", Clock.systemUTC(), time -> { });
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftChangeServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stop() throws InterruptedException {
        channel.shutdownNow();
        server.shutdownNow();
        server.awaitTermination(5, TimeUnit.SECONDS);
        fake.close();
    }

    private List<WatchChangesResponse> firstPass(String resumeCursor, SiteChangeWatch sites)
            throws Exception {
        try (Context.CancellableContext watch = Context.current().withCancellation()) {
            return watch.call(() -> {
                Iterator<WatchChangesResponse> stream = stub.watchChanges(WatchChangesRequest
                        .newBuilder().setResumeCursor(resumeCursor).setSiteChanges(sites)
                        .setPollInterval(com.google.protobuf.Duration.newBuilder().setSeconds(3600))
                        .build());
                List<WatchChangesResponse> pass = new ArrayList<>();
                while (stream.hasNext()) {
                    WatchChangesResponse next = stream.next();
                    pass.add(next);
                    if (next.hasPassCompleted()) {
                        break;
                    }
                }
                return pass;
            });
        }
    }

    private static SyncCursor cursor(WatchChangesResponse message) throws Exception {
        return SyncCursor.parseFrom(Base64.getUrlDecoder().decode(message.getResumeCursor()));
    }

    @Test
    void theFirstPassOfASiteTakesOnlyTheEndOfItsLog() throws Exception {
        fake.stub(WEB + "/web", "{\"CurrentChangeToken\":{\"StringValue\":\"web-end\"}}");

        List<WatchChangesResponse> pass = firstPass("",
                SiteChangeWatch.newBuilder().addSiteIds("s1").build());

        assertThat(fake.only(WEB + "/web").query()).isEqualTo("$select=CurrentChangeToken");
        assertThat(fake.requests()).noneMatch(request -> request.path().endsWith("/getchanges"));
        assertThat(pass).extracting(WatchChangesResponse::getEventCase).containsExactly(
                WatchChangesResponse.EventCase.SITE_CHANGE_SNAPSHOT,
                WatchChangesResponse.EventCase.PASS_COMPLETED);
        assertThat(pass.get(0).getSiteChangeSnapshot()).isEqualTo(SiteChangeSnapshot.newBuilder()
                .setSiteId("s1").setSource(ChangeSource.CHANGE_SOURCE_CRAWL).build());
        WatchChangesResponse last = pass.get(1);
        assertThat(last.getPassCompleted().getChangeCount()).isZero();
        assertThat(cursor(last).getSiteChangeTokensMap()).containsExactly(Map.entry("s1", "web-end"));
        assertThat(cursor(last).getDriveDeltaLinksMap()).containsEntry("d1", driveLink);
    }

    @Test
    void aResumedPassStreamsTheEntriesSinceTheTokenAndKeepsTheRestOfTheCursor() throws Exception {
        String users = fake.baseUrl() + "/users/delta?$deltatoken=u1";
        String resume = SyncCursors.encode(SyncCursor.newBuilder().putDriveDeltaLinks("d1", driveLink)
                .setUsersDeltaLink(users).putSiteChangeTokens("s1", "t1")
                .putSiteChangeTokens("other", "t9").build());
        fake.respond(WEB + "/site/getchanges", request -> FakeGraphServer.Stub.json(
                request.body().contains("\"t2\"") ? "{\"value\":[]}" : """
                        {"value":[{"odata.type":"SP.ChangeItem","ChangeToken":{"StringValue":"t2"},
                          "ChangeType":11,"SiteId":"%s","WebId":"%s","ListId":"%s","ItemId":4,
                          "Time":"2026-11-24T17:00:00Z"}]}""".formatted(SITE, WEB_ID, LIST_ID)));

        List<WatchChangesResponse> pass = firstPass(resume, SiteChangeWatch.newBuilder()
                .addSiteIds("s1").setScope(SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION)
                .addObjects(SiteChangeObject.SITE_CHANGE_OBJECT_ITEM)
                .addTypes(SiteChangeType.SITE_CHANGE_TYPE_ROLE_ASSIGNMENT_ADD).build());

        assertThat(syncedFrom).containsExactly(resume);
        assertThat(pass).extracting(WatchChangesResponse::getEventCase).containsExactly(
                WatchChangesResponse.EventCase.SITE_CHANGE,
                WatchChangesResponse.EventCase.SITE_CHANGE_SNAPSHOT,
                WatchChangesResponse.EventCase.PASS_COMPLETED);
        assertThat(pass).extracting(WatchChangesResponse::getSequence).containsExactly(1L, 2L, 3L);
        SiteChange change = pass.get(0).getSiteChange();
        assertThat(change.getType()).isEqualTo(SiteChangeType.SITE_CHANGE_TYPE_ROLE_ASSIGNMENT_ADD);
        assertThat(change.getSecurable().getListItem().getItemId()).isEqualTo("4");
        assertThat(pass.get(1).getSiteChangeSnapshot()).isEqualTo(SiteChangeSnapshot.newBuilder()
                .setSiteId("s1").setSource(ChangeSource.CHANGE_SOURCE_DELTA).setChangeCount(1).build());
        WatchChangesResponse last = pass.get(2);
        assertThat(last.getPassCompleted().getChangeCount()).isEqualTo(1);
        assertThat(cursor(last).getSiteChangeTokensMap())
                .containsOnly(Map.entry("s1", "t2"), Map.entry("other", "t9"));
        assertThat(cursor(last).getUsersDeltaLink()).isEqualTo(users);
        String asked = fake.requestsTo(WEB + "/site/getchanges").get(0).body();
        assertThat(Recordings.JSON.readTree(asked).path("query").path("ChangeTokenStart")
                .path("StringValue").asText()).isEqualTo("t1");
        assertThat(asked).contains("\"Item\":true", "\"RoleAssignmentAdd\":true")
                .doesNotContain("\"Group\"", "\"Update\"");
    }

    @Test
    void theLogOfAListCannotBeWatched() {
        assertThatThrownBy(() -> firstPass("", SiteChangeWatch.newBuilder().addSiteIds("s1")
                .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_LIST).build()))
                .isInstanceOfSatisfying(StatusRuntimeException.class, e -> assertThat(
                        e.getStatus().getCode()).isEqualTo(io.grpc.Status.Code.INVALID_ARGUMENT));
    }
}
