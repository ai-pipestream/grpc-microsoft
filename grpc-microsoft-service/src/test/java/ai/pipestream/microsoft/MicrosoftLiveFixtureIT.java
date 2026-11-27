package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.EnsureListRequest;
import ai.pipestream.microsoft.v1.EnsureListResponse;
import ai.pipestream.microsoft.v1.FieldValue;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListListItemsRequest;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListTemplate;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.HexFormat;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Makes sure the test site holds one small file at the root of its default
 * document library and one item in a list of its own, so that the read-only
 * coverage run finds an item to read. It writes only what is missing and
 * removes nothing.
 *
 * <p>Runs only when {@code MICROSOFT_TENANT_ID}, {@code MICROSOFT_CLIENT_ID},
 * {@code MICROSOFT_CLIENT_SECRET} and {@code MICROSOFT_SITE_ID} are set and
 * {@code MICROSOFT_LIVE_FIXTURE=1}. Run with
 * {@code ./gradlew :grpc-microsoft-service:liveFixtureTest}.</p>
 */
class MicrosoftLiveFixtureIT {

    /** Environment switch that allows writing the fixture to the tenant. */
    static final String ENV_LIVE_FIXTURE = "MICROSOFT_LIVE_FIXTURE";

    /** The file at the root of the default document library. */
    static final String FILE_NAME = "pipestream-coverage-sample.txt";

    /** The list that holds the one item. */
    static final String LIST_NAME = "pipestream-coverage";

    private static final String TAG = "[MicrosoftLiveFixtureIT] ";

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void theTestSiteHoldsOneFileAndOneListItem() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID)),
                "no live Microsoft application credentials and site in the environment; skipping");
        assumeTrue("1".equals(System.getenv(ENV_LIVE_FIXTURE)),
                ENV_LIVE_FIXTURE + " is not 1; not writing to the tenant");
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment();
        GraphAuth auth = new GraphAuth(config.authConfig());
        String token = auth.clientCredentials().accessToken();
        String graphBaseUrl = GraphLinks.stripTrailingSlashes(config.graphBaseUrl());
        GraphClient client = new GraphClient(graphBaseUrl, () -> token, AuthMode.AUTH_MODE_APPLICATION);
        String site = "/sites/" + GraphPaths.site(config.siteId());
        String siteId = client.get(site + "?$select=id").path("id").asText("");
        String driveId = config.driveIds().isEmpty()
                ? client.get(site + "/drive?$select=id").path("id").asText("")
                : config.driveIds().get(0);
        assertThat(siteId).as("the site id").isNotBlank();
        assertThat(driveId).as("the default document library").isNotBlank();

        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, new GraphFiles(client),
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES))
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub = MicrosoftServiceGrpc.newBlockingStub(channel);
            System.out.println(TAG + "file: " + file(stub, MicrosoftServiceGrpc.newStub(channel), driveId));
            System.out.println(TAG + "list item: " + listItem(stub, siteId));
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            server.awaitTermination(10, TimeUnit.SECONDS);
            executor.shutdownNow();
        }
    }

    /** Uploads the file unless the drive root already holds it. */
    private static String file(MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub,
            MicrosoftServiceGrpc.MicrosoftServiceStub async, String driveId) throws Exception {
        boolean[] held = new boolean[1];
        stub.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId).build())
                .forEachRemaining(response -> held[0] |= FILE_NAME.equals(response.getItem().getName()));
        if (held[0]) {
            return "exists";
        }
        byte[] content = "A sample file for the read-only coverage run. Safe to delete.\n"
                .getBytes(StandardCharsets.UTF_8);
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<Throwable> error = new AtomicReference<>();
        AtomicReference<UploadItemResponse> last = new AtomicReference<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(new StreamObserver<>() {
            @Override
            public void onNext(UploadItemResponse value) {
                last.set(value);
            }

            @Override
            public void onError(Throwable t) {
                error.set(t);
                done.countDown();
            }

            @Override
            public void onCompleted() {
                done.countDown();
            }
        });
        requests.onNext(UploadItemRequest.newBuilder().setHeader(UploadItemHeader.newBuilder()
                .setDriveId(driveId).setName(FILE_NAME).setMediaType("text/plain").setSize(content.length)
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_FAIL)).build());
        requests.onNext(UploadItemRequest.newBuilder().setChunk(ByteString.copyFrom(content)).build());
        requests.onNext(UploadItemRequest.newBuilder().setEnd(UploadItemEnd.newBuilder()
                .setSha256(HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content))))
                .build());
        requests.onCompleted();
        assertThat(done.await(2, TimeUnit.MINUTES)).as("the upload ended in time").isTrue();
        assertThat(error.get()).as("the upload").isNull();
        assertThat(last.get().getItem().getName()).isEqualTo(FILE_NAME);
        return "created";
    }

    /** Ensures the list and creates its item unless it already holds one. */
    private static String listItem(MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub, String siteId) {
        EnsureListResponse list = stub.ensureList(EnsureListRequest.newBuilder().setSiteId(siteId)
                .setDisplayName(LIST_NAME)
                .setDescription("One item for the read-only coverage run; safe to delete.")
                .setTemplate(ListTemplate.LIST_TEMPLATE_GENERIC_LIST).build());
        String listId = list.getList().getId();
        if (stub.listListItems(ListListItemsRequest.newBuilder().setSiteId(siteId).setListId(listId).build())
                .hasNext()) {
            return "exists (list " + list.getState() + ")";
        }
        stub.createListItem(CreateListItemRequest.newBuilder()
                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId))
                .addFields(FieldValue.newBuilder().setName("Title").setText("coverage sample")).build());
        return "created (list " + list.getState() + ")";
    }
}
