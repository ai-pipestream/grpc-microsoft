package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListChildrenResponse;
import ai.pipestream.microsoft.v1.ListDrivesRequest;
import ai.pipestream.microsoft.v1.ListSitesRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftGrpcServiceTest {

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;

    @BeforeEach
    void startStack() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(new MicrosoftGrpcService(config, files,
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES))
                .build()
                .start();
        channel = InProcessChannelBuilder.forName(name).directExecutor().build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stopStack() {
        if (channel != null) {
            channel.shutdownNow();
        }
        if (server != null) {
            server.shutdownNow();
        }
        if (fake != null) {
            fake.close();
        }
    }

    @Test
    void getMeAndListChildrenAndSync() {
        fake.stub("/me", MicrosoftFixtures.meJson());
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/children", MicrosoftFixtures.childrenJson(
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));

        assertThat(stub.getMe(GetMeRequest.getDefaultInstance()).getUser().getId())
                .isEqualTo("user-1");

        List<ListChildrenResponse> children = new ArrayList<>();
        stub.listChildren(ListChildrenRequest.newBuilder().setDriveId("drive-1").build())
                .forEachRemaining(children::add);
        assertThat(children).extracting(r -> r.getItem().getId()).containsExactly("file-1");

        List<SyncResponse> events = new ArrayList<>();
        stub.sync(SyncRequest.getDefaultInstance()).forEachRemaining(events::add);
        assertThat(events).anyMatch(SyncResponse::hasChange);
        assertThat(events).anyMatch(SyncResponse::hasSnapshot);
        assertThat(events.get(events.size() - 1).getResumeCursor()).isNotBlank();
        assertThat(fake.requests()).allMatch(r -> r.authorization().equals("Bearer token"));
    }

    @Test
    void listSitesDrivesAndGetItem() {
        fake.stub("/sites", MicrosoftFixtures.sitesJson(
                MicrosoftFixtures.siteJson("site-1", "Docs")));
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/sites/site-1/drives", MicrosoftFixtures.drivesJson(
                MicrosoftFixtures.driveJson("drive-2", "Library")));
        fake.stub("/drives/drive-1/items/file-1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1"));

        assertThat(stub.listSites(ListSitesRequest.newBuilder().setLimit(1).build())
                .getSitesList()).extracting(s -> s.getId()).containsExactly("site-1");
        assertThat(stub.listDrives(ListDrivesRequest.getDefaultInstance())
                .getDrivesList()).extracting(d -> d.getId()).containsExactly("drive-1");
        assertThat(stub.listDrives(ListDrivesRequest.newBuilder().setSiteId("site-1").build())
                .getDrivesList()).extracting(d -> d.getId()).containsExactly("drive-2");
        assertThat(stub.getItem(GetItemRequest.newBuilder()
                .setDriveId("drive-1")
                .setItemId("file-1")
                .build()).getItem().getName()).isEqualTo("notes.txt");
    }
}
