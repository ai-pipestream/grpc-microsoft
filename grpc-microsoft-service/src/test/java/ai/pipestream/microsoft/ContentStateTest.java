package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ContentState;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
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

import java.nio.charset.StandardCharsets;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A Sync that was asked for content says, on every drive item, what became
 * of its bytes. An item without {@code content} is otherwise the same
 * whether the file is over the inline limit, empty, a folder or flagged as
 * malware, and a consumer could only guess by re-deriving the server's
 * limit.
 */
class ContentStateTest {

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
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token",
                AuthMode.AUTH_MODE_UNSPECIFIED));
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, files, 0, null, null, 4))
                .build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);

        fake.stub("/drives/d1", MicrosoftFixtures.driveJson("d1", "Docs"));
        fake.stub("/drives/d1/root/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/d1/root/delta?token=t1",
                file("small", 5),
                file("big", 40),
                file("grew", 5),
                file("empty", 0),
                MicrosoftFixtures.folderJson("folder", "Reports", "d1"),
                file("notebook", 40).replace("\"file\": {\"mimeType\": \"text/plain\"}",
                        "\"package\": {\"type\": \"oneNote\"}"),
                file("infected", 5).replace("\"file\":", "\"malware\": {\"description\": \"x\"}, \"file\":")));
        fake.stub("/drives/d1/items/small/content", bytes(5));
        fake.stub("/drives/d1/items/big/content", bytes(40));
        // Listed as 5 bytes, but the download is 40.
        fake.stub("/drives/d1/items/grew/content", bytes(40));
        fake.stub("/drives/d1/items/infected/content", bytes(5));
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    @Test
    void everyItemOfASyncThatAskedForContentSaysWhatBecameOfItsBytes() {
        Map<String, DriveItem> items = sync(SyncRequest.newBuilder().addDriveIds("d1")
                .setIncludeContent(true).setMaxInlineBytes(10).build());

        assertThat(items.get("small").getContent().size()).isEqualTo(5);
        assertThat(items.get("small").getContentState()).isEqualTo(ContentState.CONTENT_STATE_INLINED);
        assertThat(items.get("big").hasContent()).isFalse();
        assertThat(items.get("big").getContentState())
                .as("a file listed as larger than the inline limit")
                .isEqualTo(ContentState.CONTENT_STATE_OMITTED_OVER_LIMIT);
        assertThat(items.get("grew").hasContent()).isFalse();
        assertThat(items.get("grew").getContentState())
                .as("a file whose download is larger than the inline limit")
                .isEqualTo(ContentState.CONTENT_STATE_OMITTED_OVER_LIMIT);
        assertThat(items.get("empty").hasContent()).isFalse();
        assertThat(items.get("empty").getContentState()).isEqualTo(ContentState.CONTENT_STATE_EMPTY);
        assertThat(items.get("folder").getContentState())
                .isEqualTo(ContentState.CONTENT_STATE_OMITTED_NOT_A_FILE);
        assertThat(items.get("notebook").getContentState())
                .isEqualTo(ContentState.CONTENT_STATE_OMITTED_NOT_A_FILE);
        assertThat(items.get("infected").hasContent()).isFalse();
        assertThat(items.get("infected").getContentState())
                .isEqualTo(ContentState.CONTENT_STATE_OMITTED_MALWARE);
        // Nothing is downloaded that is known not to be sent.
        assertThat(fake.requestsTo("/drives/d1/items/big/content")).isEmpty();
        assertThat(fake.requestsTo("/drives/d1/items/infected/content")).isEmpty();
    }

    @Test
    void aSyncThatDidNotAskForContentSaysNothingAboutIt() {
        Map<String, DriveItem> items = sync(SyncRequest.newBuilder().addDriveIds("d1").build());

        assertThat(items.values()).allSatisfy(item -> {
            assertThat(item.hasContent()).isFalse();
            assertThat(item.getContentState()).isEqualTo(ContentState.CONTENT_STATE_UNSPECIFIED);
        });
        assertThat(fake.requestsTo("/drives/d1/items/small/content")).isEmpty();
    }

    private Map<String, DriveItem> sync(SyncRequest request) {
        Map<String, DriveItem> items = new LinkedHashMap<>();
        stub.sync(request).forEachRemaining(response -> {
            if (response.getEventCase() == SyncResponse.EventCase.CHANGE && response.getChange()
                    .getEntity().getEntityCase() == MicrosoftEntity.EntityCase.DRIVE_ITEM) {
                DriveItem item = response.getChange().getEntity().getDriveItem();
                items.put(item.getId(), item);
            }
        });
        return items;
    }

    private static String file(String id, long size) {
        return MicrosoftFixtures.fileJson(id, id + ".txt", "d1")
                .replace("\"size\": 12", "\"size\": " + size);
    }

    private static FakeGraphServer.Stub bytes(int count) {
        return FakeGraphServer.Stub.bytes("x".repeat(count).getBytes(StandardCharsets.US_ASCII),
                "text/plain");
    }
}
