package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DirectoryKind;
import ai.pipestream.microsoft.v1.DirectorySnapshot;
import ai.pipestream.microsoft.v1.DirectoryWatch;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.microsoft.v1.WatchChangesRequest;
import ai.pipestream.microsoft.v1.WatchChangesResponse;
import io.grpc.Context;
import io.grpc.ManagedChannel;
import io.grpc.Server;
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

/**
 * The Entra part of a WatchChanges pass against the in-process fake Graph:
 * what it asks for, what it streams and what the resume cursor carries. The
 * live run and its replay cover the same pass with real answers.
 */
class MicrosoftChangeDirectoryTest {

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

    private List<WatchChangesResponse> firstPass(String resumeCursor, DirectoryWatch directory)
            throws Exception {
        try (Context.CancellableContext watch = Context.current().withCancellation()) {
            return watch.call(() -> {
                Iterator<WatchChangesResponse> stream = stub.watchChanges(WatchChangesRequest
                        .newBuilder().setResumeCursor(resumeCursor).setDirectory(directory)
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

    private static DirectorySnapshot snapshot(List<WatchChangesResponse> pass, DirectoryKind kind) {
        return pass.stream().filter(WatchChangesResponse::hasDirectorySnapshot)
                .map(WatchChangesResponse::getDirectorySnapshot)
                .filter(one -> one.getKind() == kind).findFirst().orElseThrow();
    }

    @Test
    void theFirstPassListsEveryUserAndGroupAndTheCursorCarriesTheirLinks() throws Exception {
        String users = fake.baseUrl() + "/users/delta?$deltatoken=u1";
        String groups = fake.baseUrl() + "/groups/delta?$deltatoken=g1";
        fake.stub("/users/delta", MicrosoftFixtures.deltaJson(users,
                "{\"id\":\"u1\",\"userPrincipalName\":\"ada@contoso.example\"}",
                "{\"id\":\"u2\",\"userPrincipalName\":\"bob@contoso.example\"}"));
        fake.stub("/groups/delta", MicrosoftFixtures.deltaJson(groups, """
                {"id":"g1","displayName":"Research","members@delta":[
                  {"@odata.type":"#microsoft.graph.user","id":"u1"}]}"""));

        List<WatchChangesResponse> pass = firstPass("", DirectoryWatch.newBuilder().setUsers(true)
                .setGroups(true).setGroupMembers(true).build());

        assertThat(fake.only("/users/delta").query()).isEqualTo("$select=" + GraphPaths.USER_SELECT);
        assertThat(fake.only("/groups/delta").query())
                .isEqualTo("$select=" + GraphPaths.GROUP_SELECT + ",members");
        assertThat(pass).extracting(WatchChangesResponse::getEventCase).containsExactly(
                WatchChangesResponse.EventCase.USER_CHANGE, WatchChangesResponse.EventCase.USER_CHANGE,
                WatchChangesResponse.EventCase.DIRECTORY_SNAPSHOT,
                WatchChangesResponse.EventCase.GROUP_CHANGE,
                WatchChangesResponse.EventCase.DIRECTORY_SNAPSHOT,
                WatchChangesResponse.EventCase.PASS_COMPLETED);
        assertThat(pass).extracting(WatchChangesResponse::getSequence)
                .containsExactly(1L, 2L, 3L, 4L, 5L, 6L);
        assertThat(pass.get(3).getGroupChange().getMemberChanges(0).getId()).isEqualTo("u1");
        DirectorySnapshot userSnapshot = snapshot(pass, DirectoryKind.DIRECTORY_KIND_USERS);
        assertThat(userSnapshot.getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_CRAWL);
        assertThat(userSnapshot.getChangeCount()).isEqualTo(2);
        assertThat(snapshot(pass, DirectoryKind.DIRECTORY_KIND_GROUPS).getChangeCount()).isEqualTo(1);
        WatchChangesResponse last = pass.get(pass.size() - 1);
        assertThat(last.getPassCompleted().getChangeCount()).isEqualTo(3);
        assertThat(cursor(last).getUsersDeltaLink()).isEqualTo(users);
        assertThat(cursor(last).getGroupsDeltaLink()).isEqualTo(groups);
        assertThat(cursor(last).getDriveDeltaLinksMap()).containsEntry("d1", driveLink);
    }

    @Test
    void aResumedPassReadsOnlyTheChangesAndARejectedLinkListsEverythingAgain() throws Exception {
        String users = fake.baseUrl() + "/users/delta?$deltatoken=u1";
        String groups = fake.baseUrl() + "/groups/delta?$deltatoken=g1";
        String resume = SyncCursors.encode(SyncCursor.newBuilder().putDriveDeltaLinks("d1", driveLink)
                .setUsersDeltaLink(users).setGroupsDeltaLink(groups).build());
        fake.stub("/users/delta?$deltatoken=u1", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/users/delta?$deltatoken=u2",
                "{\"id\":\"u2\",\"@removed\":{\"reason\":\"changed\"}}"));
        fake.stub("/groups/delta?$deltatoken=g1", FakeGraphServer.error(410, "syncStateNotFound",
                Map.of("Location", fake.baseUrl() + "/groups/delta?token=again")));
        fake.stub("/groups/delta?token=again", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/groups/delta?$deltatoken=g2", "{\"id\":\"g1\"}", "{\"id\":\"g2\"}"));

        List<WatchChangesResponse> pass = firstPass(resume, DirectoryWatch.newBuilder().setUsers(true)
                .setGroups(true).build());

        assertThat(syncedFrom).containsExactly(resume);
        assertThat(pass.get(0).getUserChange().getRemoved()).isTrue();
        assertThat(snapshot(pass, DirectoryKind.DIRECTORY_KIND_USERS).getSource())
                .isEqualTo(ChangeSource.CHANGE_SOURCE_DELTA);
        DirectorySnapshot groupSnapshot = snapshot(pass, DirectoryKind.DIRECTORY_KIND_GROUPS);
        assertThat(groupSnapshot.getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_CRAWL);
        assertThat(groupSnapshot.getChangeCount()).isEqualTo(2);
        WatchChangesResponse last = pass.get(pass.size() - 1);
        assertThat(cursor(last).getUsersDeltaLink()).endsWith("$deltatoken=u2");
        assertThat(cursor(last).getGroupsDeltaLink()).endsWith("$deltatoken=g2");
    }

    @Test
    void aPartOfTheDirectoryThatIsNotWatchedKeepsItsLink() throws Exception {
        String users = fake.baseUrl() + "/users/delta?$deltatoken=u1";
        String groups = fake.baseUrl() + "/groups/delta?$deltatoken=g1";
        String resume = SyncCursors.encode(SyncCursor.newBuilder().setUsersDeltaLink(users)
                .setGroupsDeltaLink(groups).build());
        fake.stub("/users/delta?$deltatoken=u1", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/users/delta?$deltatoken=u2"));

        List<WatchChangesResponse> pass = firstPass(resume, DirectoryWatch.newBuilder().setUsers(true)
                .build());

        assertThat(fake.requestsTo("/groups/delta")).isEmpty();
        assertThat(pass).extracting(WatchChangesResponse::getEventCase).containsExactly(
                WatchChangesResponse.EventCase.DIRECTORY_SNAPSHOT,
                WatchChangesResponse.EventCase.PASS_COMPLETED);
        WatchChangesResponse last = pass.get(pass.size() - 1);
        assertThat(cursor(last).getUsersDeltaLink()).endsWith("$deltatoken=u2");
        assertThat(cursor(last).getGroupsDeltaLink()).isEqualTo(groups);
    }
}
