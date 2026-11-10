package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ClearValue;
import ai.pipestream.microsoft.v1.ColumnOutcome;
import ai.pipestream.microsoft.v1.ColumnScope;
import ai.pipestream.microsoft.v1.ColumnSpec;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CopyDriveItemRequest;
import ai.pipestream.microsoft.v1.CopyDriveItemResponse;
import ai.pipestream.microsoft.v1.CreateFolderRequest;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeleteListItemRequest;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.EnsureColumnsRequest;
import ai.pipestream.microsoft.v1.EnsureListRequest;
import ai.pipestream.microsoft.v1.EnsureListResponse;
import ai.pipestream.microsoft.v1.EnsureState;
import ai.pipestream.microsoft.v1.FieldOutcome;
import ai.pipestream.microsoft.v1.FieldValue;
import ai.pipestream.microsoft.v1.FieldWriteState;
import ai.pipestream.microsoft.v1.FileSystemInfo;
import ai.pipestream.microsoft.v1.HyperlinkValue;
import ai.pipestream.microsoft.v1.ListItemPatch;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemTarget;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListTemplate;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MultiChoiceValue;
import ai.pipestream.microsoft.v1.MultiLookupValue;
import ai.pipestream.microsoft.v1.MultiPersonValue;
import ai.pipestream.microsoft.v1.MultiTermValue;
import ai.pipestream.microsoft.v1.MutateRequest;
import ai.pipestream.microsoft.v1.MutateResponse;
import ai.pipestream.microsoft.v1.PersonValue;
import ai.pipestream.microsoft.v1.RestoreDriveItemVersionRequest;
import ai.pipestream.microsoft.v1.TermValue;
import ai.pipestream.microsoft.v1.UpdateDriveItemRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsResponse;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.protobuf.ByteString;
import com.google.protobuf.Duration;
import com.google.protobuf.Timestamp;
import com.google.rpc.RetryInfo;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.protobuf.StatusProto;
import io.grpc.stub.StreamObserver;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.catchThrowableOfType;

/** The write rpcs of {@code MicrosoftService} against the in-process fake Graph. */
class MicrosoftWriteRpcTest {

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final int UNIT = GraphFiles.FRAGMENT_UNIT_BYTES;
    private static final String SITE = "contoso.sharepoint.com,c1,w1";
    private static final String LIST_ROUTE = "/sites/" + SITE + "/lists/l1";

    private FakeGraphServer fake;
    private Server server;
    private ManagedChannel channel;
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private MicrosoftServiceGrpc.MicrosoftServiceStub async;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String name = InProcessServerBuilder.generateName();
        server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, files, 0)).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        stub = MicrosoftServiceGrpc.newBlockingStub(channel);
        async = MicrosoftServiceGrpc.newStub(channel);
    }

    @AfterEach
    void stop() {
        channel.shutdownNow();
        server.shutdownNow();
        executor.shutdownNow();
        fake.close();
    }

    // ------------------------------------------------------------ helpers

    /** Collects the responses of a streaming call. */
    private static final class Collector<T> implements StreamObserver<T> {
        private final BlockingQueue<T> messages = new LinkedBlockingQueue<>();
        private final List<T> all = new ArrayList<>();
        private final CountDownLatch done = new CountDownLatch(1);
        private final AtomicReference<Throwable> error = new AtomicReference<>();

        @Override
        public synchronized void onNext(T value) {
            all.add(value);
            messages.add(value);
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

        T next() throws InterruptedException {
            T value = messages.poll(20, TimeUnit.SECONDS);
            assertThat(value).as("a response within 20 s").isNotNull();
            return value;
        }

        synchronized List<T> finished() throws InterruptedException {
            return List.copyOf(all);
        }

        List<T> await() throws InterruptedException {
            assertThat(done.await(30, TimeUnit.SECONDS)).as("the call ended").isTrue();
            assertThat(error.get()).isNull();
            return finished();
        }

        Status.Code awaitError() throws InterruptedException {
            assertThat(done.await(30, TimeUnit.SECONDS)).as("the call ended").isTrue();
            assertThat(error.get()).isNotNull();
            return Status.fromThrowable(error.get()).getCode();
        }
    }

    private static byte[] bytes(int length) {
        byte[] content = new byte[length];
        new Random(length).nextBytes(content);
        return content;
    }

    private static String sha256(byte[] content) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
    }

    private static UploadItemRequest header(UploadItemHeader.Builder header) {
        return UploadItemRequest.newBuilder().setHeader(header).build();
    }

    private static UploadItemRequest end(String sha256) {
        return UploadItemRequest.newBuilder()
                .setEnd(UploadItemEnd.newBuilder().setSha256(sha256)).build();
    }

    private static void sendChunks(StreamObserver<UploadItemRequest> requests, byte[] content,
            int from, int to, int chunk) {
        for (int offset = from; offset < to; offset += chunk) {
            requests.onNext(UploadItemRequest.newBuilder()
                    .setChunk(ByteString.copyFrom(content, offset, Math.min(chunk, to - offset)))
                    .build());
        }
    }

    private Collector<UploadItemResponse> upload(UploadItemHeader.Builder header, byte[] content,
            int chunk, String sha256) {
        Collector<UploadItemResponse> responses = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(responses);
        requests.onNext(header(header));
        sendChunks(requests, content, 0, content.length, chunk);
        requests.onNext(end(sha256));
        requests.onCompleted();
        return responses;
    }

    private String sessionAt(String createRoute, String itemJson, ByteArrayOutputStream[] stored) {
        String uploadUrl = fake.baseUrl() + "/upload/s1";
        fake.stub(createRoute, "{\"uploadUrl\":\"" + uploadUrl + "\"}");
        stored[0] = fake.uploadSession("/upload/s1", itemJson);
        return uploadUrl;
    }

    private JsonNode body(FakeGraphServer.RecordedRequest request) throws Exception {
        return JSON.readTree(request.body());
    }

    private static DriveItemRef item(String id) {
        return DriveItemRef.newBuilder().setDriveId("d1").setItemId(id).build();
    }

    private static Status.Code code(Runnable call) {
        StatusRuntimeException e = catchThrowableOfType(StatusRuntimeException.class, call::run);
        return e == null ? Status.Code.OK : e.getStatus().getCode();
    }

    // ------------------------------------------------------------ upload

    @Test
    void uploadItemOfSmallKnownSizeUsesOnePutAndSetsColumns() throws Exception {
        byte[] content = bytes(1000);
        fake.stub("PUT", "/drives/d1/items/p1:/a b.txt:/content",
                FakeGraphServer.Stub.json(MicrosoftFixtures.fileJson("f1", "a b.txt", "d1")));
        fake.stub("/drives/d1/items/f1/listItem/fields", "{\"Summary\":\"short\"}");
        fake.stub("/drives/d1/items/f1", MicrosoftFixtures.fileJson("f1", "a b.txt", "d1"));

        List<UploadItemResponse> responses = upload(UploadItemHeader.newBuilder()
                .setDriveId("d1").setParentId("p1").setName("a b.txt").setMediaType("text/plain")
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_REPLACE)
                .setExpectedEtag("\"e1\"").setSize(content.length)
                .addFields(FieldValue.newBuilder().setName("Summary").setNote("short")),
                content, 300, sha256(content)).await();

        assertThat(responses).hasSize(1);
        DriveItem item = responses.get(0).getItem();
        assertThat(item.getId()).isEqualTo("f1");
        assertThat(item.getListColumnsList()).extracting(c -> c.getName() + "=" + c.getStringValue())
                .containsExactly("Summary=short");
        FakeGraphServer.RecordedRequest put = fake.requests().stream()
                .filter(r -> r.method().equals("PUT")).findFirst().orElseThrow();
        assertThat(put.path()).isEqualTo("/drives/d1/items/p1:/a b.txt:/content");
        assertThat(put.query()).isEqualTo("@microsoft.graph.conflictBehavior=replace");
        assertThat(put.ifMatch()).isEqualTo("\"e1\"");
        assertThat(put.bodyLength()).isEqualTo(1000);
        FakeGraphServer.RecordedRequest patch = fake.requests().stream()
                .filter(r -> r.method().equals("PATCH")).findFirst().orElseThrow();
        assertThat(body(patch).path("Summary").asText()).isEqualTo("short");
    }

    @Test
    void uploadItemOfLargeKnownSizeSendsFragmentsThroughAnUploadSession() throws Exception {
        byte[] content = bytes(13 * UNIT + 12345);
        ByteArrayOutputStream[] stored = new ByteArrayOutputStream[1];
        String createRoute = "/drives/d1/root:/docs/big.bin:/createUploadSession";
        String uploadUrl = sessionAt(createRoute,
                MicrosoftFixtures.fileJson("big", "big.bin", "d1"), stored);

        List<UploadItemResponse> responses = upload(UploadItemHeader.newBuilder()
                .setDriveId("d1").setParentPath("/docs").setName("big.bin")
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_RENAME)
                .setExpectedEtag("\"e9\"").setSize(content.length).setFragmentBytes(5 * UNIT + 5),
                content, 100_000, sha256(content)).await();

        assertThat(stored[0].toByteArray()).isEqualTo(content);
        assertThat(responses.get(responses.size() - 1).getItem().getId()).isEqualTo("big");
        assertThat(responses.subList(0, responses.size() - 1))
                .allMatch(UploadItemResponse::hasProgress)
                .allMatch(r -> !r.getProgress().getResumeToken().isEmpty())
                .extracting(r -> r.getProgress().getCommittedBytes())
                .containsExactly(0L, 5L * UNIT, 10L * UNIT, (long) content.length);
        assertThat(responses.get(0).getProgress().getResumeToken()).doesNotContain("upload/s1");

        FakeGraphServer.RecordedRequest create = fake.only(createRoute);
        assertThat(create.authorization()).isEqualTo("Bearer token");
        assertThat(create.ifMatch()).isEqualTo("\"e9\"");
        assertThat(body(create).path("item").path("@microsoft.graph.conflictBehavior").asText())
                .isEqualTo("rename");
        List<FakeGraphServer.RecordedRequest> puts = fake.requestsTo("/upload/s1");
        assertThat(puts).allMatch(r -> r.authorization() == null);
        assertThat(puts).extracting(FakeGraphServer.RecordedRequest::contentRange).containsExactly(
                "bytes 0-" + (5 * UNIT - 1) + "/" + content.length,
                "bytes " + 5 * UNIT + "-" + (10 * UNIT - 1) + "/" + content.length,
                "bytes " + 10 * UNIT + "-" + (content.length - 1) + "/" + content.length);
        assertThat(uploadUrl).endsWith("/upload/s1");
    }

    @Test
    void uploadItemOfUnknownSizeIsSpooledThenSent() throws Exception {
        byte[] small = bytes(2000);
        List<UploadItemResponse> responses = upload(UploadItemHeader.newBuilder()
                .setDriveId("d1").setName("small.bin"), small, 700, "").await();
        assertThat(responses).hasSize(1);
        assertThat(responses.get(0).getItem().getId()).isEqualTo("uploaded");
        FakeGraphServer.RecordedRequest put = fake.only("/drives/d1/root:/small.bin:/content");
        assertThat(put.bodyLength()).isEqualTo(2000);
        assertThat(put.query()).isEmpty();
        assertThat(put.ifMatch()).isNull();

        byte[] large = bytes(GraphFiles.SIMPLE_UPLOAD_LIMIT + 77);
        ByteArrayOutputStream[] stored = new ByteArrayOutputStream[1];
        sessionAt("/drives/d1/root:/large.bin:/createUploadSession",
                MicrosoftFixtures.fileJson("large", "large.bin", "d1"), stored);
        responses = upload(UploadItemHeader.newBuilder().setDriveId("d1").setName("large.bin"),
                large, 1 << 20, sha256(large)).await();
        assertThat(stored[0].toByteArray()).isEqualTo(large);
        assertThat(responses.get(responses.size() - 1).getItem().getId()).isEqualTo("large");
        assertThat(fake.requestsTo("/upload/s1")).hasSize(1);
    }

    @Test
    void uploadItemResumesABrokenUploadFromItsToken() throws Exception {
        byte[] content = bytes(GraphFiles.SIMPLE_UPLOAD_LIMIT + 3 * UNIT);
        ByteArrayOutputStream[] stored = new ByteArrayOutputStream[1];
        sessionAt("/drives/d1/root:/big.bin:/createUploadSession",
                MicrosoftFixtures.fileJson("big", "big.bin", "d1"), stored);
        UploadItemHeader.Builder header = UploadItemHeader.newBuilder().setDriveId("d1")
                .setName("big.bin").setSize(content.length).setFragmentBytes(4 * UNIT);

        Collector<UploadItemResponse> first = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(first);
        requests.onNext(header(header));
        int sent = 9 * UNIT;
        sendChunks(requests, content, 0, sent, 250_000);
        String token = "";
        long committed = 0;
        while (committed < 8L * UNIT) {
            UploadItemResponse progress = first.next();
            token = progress.getProgress().getResumeToken();
            committed = progress.getProgress().getCommittedBytes();
        }
        requests.onError(Status.CANCELLED.asRuntimeException());
        first.awaitError();
        assertThat(stored[0].size()).isEqualTo(8 * UNIT);

        Collector<UploadItemResponse> wrongOffset = new Collector<>();
        StreamObserver<UploadItemRequest> wrong = async.uploadItem(wrongOffset);
        wrong.onNext(header(header.clone().setResumeToken(token).setResumeOffset(UNIT)));
        assertThat(wrongOffset.awaitError()).isEqualTo(Status.Code.FAILED_PRECONDITION);

        Collector<UploadItemResponse> forged = new Collector<>();
        StreamObserver<UploadItemRequest> forgedCall = async.uploadItem(forged);
        forgedCall.onNext(header(header.clone().setResumeToken(token + "x")
                .setResumeOffset(committed)));
        assertThat(forged.awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);

        Collector<UploadItemResponse> second = new Collector<>();
        StreamObserver<UploadItemRequest> resumed = async.uploadItem(second);
        resumed.onNext(header(header.clone().setResumeToken(token).setResumeOffset(committed)));
        sendChunks(resumed, content, (int) committed, content.length, 250_000);
        resumed.onNext(end(""));
        resumed.onCompleted();
        List<UploadItemResponse> responses = second.await();

        assertThat(stored[0].toByteArray()).isEqualTo(content);
        assertThat(responses.get(0).getProgress().getCommittedBytes()).isEqualTo(committed);
        assertThat(responses.get(responses.size() - 1).getItem().getId()).isEqualTo("big");
        assertThat(fake.requestsTo("/drives/d1/root:/big.bin:/createUploadSession")).hasSize(1);
    }

    @Test
    void uploadItemRejectsAWrongDigestAndCancelsTheSession() throws Exception {
        byte[] content = bytes(GraphFiles.SIMPLE_UPLOAD_LIMIT + 10);
        ByteArrayOutputStream[] stored = new ByteArrayOutputStream[1];
        sessionAt("/drives/d1/root:/big.bin:/createUploadSession",
                MicrosoftFixtures.fileJson("big", "big.bin", "d1"), stored);

        Collector<UploadItemResponse> responses = upload(UploadItemHeader.newBuilder()
                .setDriveId("d1").setName("big.bin").setSize(content.length)
                .setFragmentBytes(GraphFiles.MAX_FRAGMENT_BYTES), content, 1 << 20, "00ff");

        assertThat(responses.awaitError()).isEqualTo(Status.Code.DATA_LOSS);
        assertThat(stored[0].size()).isZero();
        assertThat(fake.requestsTo("/upload/s1"))
                .extracting(FakeGraphServer.RecordedRequest::method).containsExactly("DELETE");
    }

    @Test
    void uploadItemValidatesItsFrames() throws Exception {
        Collector<UploadItemResponse> noHeader = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(noHeader);
        requests.onNext(UploadItemRequest.newBuilder().setChunk(ByteString.copyFromUtf8("x")).build());
        assertThat(noHeader.awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);

        assertThat(upload(UploadItemHeader.newBuilder().setDriveId("d1").setName("a/b"),
                bytes(1), 1, "").awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(upload(UploadItemHeader.newBuilder().setDriveId("d1").setName("a").setSize(5),
                bytes(6), 6, "").awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(upload(UploadItemHeader.newBuilder().setDriveId("d1").setName("a").setSize(5),
                bytes(4), 4, "").awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);

        Collector<UploadItemResponse> noEnd = new Collector<>();
        requests = async.uploadItem(noEnd);
        requests.onNext(header(UploadItemHeader.newBuilder().setDriveId("d1").setName("a")));
        requests.onCompleted();
        assertThat(noEnd.awaitError()).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(fake.requests()).isEmpty();
    }

    @Test
    void fragmentSizeIsAMultipleOf320KiBUpTo60MiB() {
        assertThat(UploadHandler.fragmentBytes(0)).isEqualTo(GraphFiles.DEFAULT_FRAGMENT_BYTES);
        assertThat(UploadHandler.fragmentBytes(1)).isEqualTo(UNIT);
        assertThat(UploadHandler.fragmentBytes(2 * UNIT + 9)).isEqualTo(2 * UNIT);
        assertThat(UploadHandler.fragmentBytes(Integer.MAX_VALUE))
                .isEqualTo(GraphFiles.MAX_FRAGMENT_BYTES)
                .matches(bytes -> bytes % UNIT == 0);
        assertThat(GraphFiles.DEFAULT_FRAGMENT_BYTES % UNIT).isZero();
    }

    // ------------------------------------------------------------ errors

    @Test
    void etagMismatchIsAbortedAndThrottlingCarriesRetryInfo() {
        fake.stub("DELETE", "/drives/d1/items/stale",
                FakeGraphServer.error(412, "preconditionFailed", Map.of()));
        assertThat(code(() -> stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder()
                .setItem(item("stale")).setExpectedEtag("\"old\"").build())))
                .isEqualTo(Status.Code.ABORTED);
        assertThat(fake.only("/drives/d1/items/stale").ifMatch()).isEqualTo("\"old\"");

        for (int i = 0; i < 3; i++) {
            fake.stubOnce("/drives/d1/items/busy",
                    FakeGraphServer.error(429, "tooManyRequests", Map.of("Retry-After", "0")));
        }
        fake.stub("/drives/d1/items/busy",
                FakeGraphServer.error(429, "tooManyRequests", Map.of("Retry-After", "7")));
        StatusRuntimeException throttled = catchThrowableOfType(StatusRuntimeException.class,
                () -> stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder()
                        .setItem(item("busy")).build()));
        assertThat(throttled.getStatus().getCode()).isEqualTo(Status.Code.RESOURCE_EXHAUSTED);
        assertThat(retryDelaySeconds(throttled)).isEqualTo(7);

        fake.stub("/drives/d1/items/down",
                FakeGraphServer.error(503, "serviceNotAvailable", Map.of("Retry-After", "0")));
        StatusRuntimeException unavailable = catchThrowableOfType(StatusRuntimeException.class,
                () -> stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder()
                        .setItem(item("down")).build()));
        assertThat(unavailable.getStatus().getCode()).isEqualTo(Status.Code.UNAVAILABLE);
        assertThat(retryDelaySeconds(unavailable)).isZero();
    }

    private static long retryDelaySeconds(StatusRuntimeException e) {
        try {
            return StatusProto.fromThrowable(e).getDetails(0).unpack(RetryInfo.class)
                    .getRetryDelay().getSeconds();
        } catch (com.google.protobuf.InvalidProtocolBufferException invalid) {
            throw new AssertionError(invalid);
        }
    }

    // ------------------------------------------------------------ drive items

    @Test
    void createFolderPostsTheFolderFacet() throws Exception {
        fake.stub("/drives/d1/items/p1/children", MicrosoftFixtures.folderJson("n1", "New", "d1"));
        fake.stub("/drives/d1/root:/a/b:/children", MicrosoftFixtures.folderJson("n2", "New", "d1"));

        assertThat(stub.createFolder(CreateFolderRequest.newBuilder().setDriveId("d1")
                .setParentId("p1").setName("New")
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_FAIL).build())
                .getItem().getId()).isEqualTo("n1");
        FakeGraphServer.RecordedRequest post = fake.only("/drives/d1/items/p1/children");
        assertThat(post.method()).isEqualTo("POST");
        assertThat(body(post).path("name").asText()).isEqualTo("New");
        assertThat(body(post).path("folder").isObject()).isTrue();
        assertThat(body(post).path("@microsoft.graph.conflictBehavior").asText()).isEqualTo("fail");

        assertThat(stub.createFolder(CreateFolderRequest.newBuilder().setDriveId("d1")
                .setParentPath("/a/b").setName("New").build()).getItem().getId()).isEqualTo("n2");
        assertThat(body(fake.only("/drives/d1/root:/a/b:/children"))
                .has("@microsoft.graph.conflictBehavior")).isFalse();
        assertThat(code(() -> stub.createFolder(CreateFolderRequest.newBuilder().setDriveId("d1")
                .setName("a/b").build()))).isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void updateDriveItemPatchesOnlyWhatIsSet() throws Exception {
        fake.stub("/drives/d1/items/f1", MicrosoftFixtures.fileJson("f1", "new.txt", "d1"));

        assertThat(stub.updateDriveItem(UpdateDriveItemRequest.newBuilder().setItem(item("f1"))
                .setExpectedEtag("\"e1\"").setName("new.txt").setNewParentId("p2")
                .setDescription("")
                .setFileSystemInfo(FileSystemInfo.newBuilder()
                        .setLastModifiedAt(Timestamp.newBuilder().setSeconds(1_700_000_000)))
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_RENAME).build())
                .getItem().getName()).isEqualTo("new.txt");

        FakeGraphServer.RecordedRequest patch = fake.only("/drives/d1/items/f1");
        assertThat(patch.method()).isEqualTo("PATCH");
        assertThat(patch.ifMatch()).isEqualTo("\"e1\"");
        assertThat(patch.query()).isEqualTo("@microsoft.graph.conflictBehavior=rename");
        JsonNode body = body(patch);
        assertThat(body.path("name").asText()).isEqualTo("new.txt");
        assertThat(body.path("parentReference").path("id").asText()).isEqualTo("p2");
        assertThat(body.path("description").asText("unset")).isEmpty();
        assertThat(body.path("fileSystemInfo").path("lastModifiedDateTime").asText())
                .isEqualTo("2023-11-14T22:13:20Z");
        assertThat(body.path("fileSystemInfo").has("createdDateTime")).isFalse();

        assertThat(code(() -> stub.updateDriveItem(UpdateDriveItemRequest.newBuilder()
                .setItem(item("f1")).build()))).isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void deleteDriveItemRecyclesOrDeletesForGood() {
        fake.stub("DELETE", "/drives/d1/items/f1", FakeGraphServer.empty(204, Map.of()));
        fake.stub("POST", "/drives/d1/items/f2/permanentDelete",
                FakeGraphServer.empty(204, Map.of()));

        assertThat(stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(item("f1"))
                .setExpectedEtag("\"e1\"").build()).getItem()).isEqualTo(item("f1"));
        assertThat(fake.only("/drives/d1/items/f1").method()).isEqualTo("DELETE");
        assertThat(fake.only("/drives/d1/items/f1").ifMatch()).isEqualTo("\"e1\"");

        assertThat(stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(item("f2"))
                .setPermanent(true).build()).getItem()).isEqualTo(item("f2"));
        assertThat(fake.only("/drives/d1/items/f2/permanentDelete").method()).isEqualTo("POST");
        assertThat(fake.requestsTo("/drives/d1/items/f2")).isEmpty();
    }

    @Test
    void restoreDriveItemVersionPostsRestoreVersion() {
        fake.stub("POST", "/drives/d1/items/f1/versions/2.0/restoreVersion",
                FakeGraphServer.empty(204, Map.of()));
        assertThat(stub.restoreDriveItemVersion(RestoreDriveItemVersionRequest.newBuilder()
                .setItem(item("f1")).setVersionId("2.0").setExpectedEtag("\"e1\"").build())
                .getVersionId()).isEqualTo("2.0");
        assertThat(fake.only("/drives/d1/items/f1/versions/2.0/restoreVersion").ifMatch())
                .isEqualTo("\"e1\"");
    }

    @Test
    void copyDriveItemStreamsProgressThenTheNewItem() throws Exception {
        fake.stub("POST", "/drives/d1/items/f1/copy", FakeGraphServer.empty(202,
                Map.of("Location", fake.baseUrl() + "/monitor/1")));
        fake.stubOnce("/monitor/1", FakeGraphServer.json(202,
                "{\"status\":\"inProgress\",\"percentageComplete\":40.5}", Map.of()));
        fake.stub("/monitor/1", "{\"status\":\"completed\",\"percentageComplete\":100,"
                + "\"resourceId\":\"copy1\"}");
        fake.stub("/drives/d2/items/copy1", MicrosoftFixtures.fileJson("copy1", "b.txt", "d2"));

        List<CopyDriveItemResponse> events = new ArrayList<>();
        stub.copyDriveItem(CopyDriveItemRequest.newBuilder().setItem(item("f1"))
                .setDestinationDriveId("d2").setDestinationParentId("p9").setName("b.txt")
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_REPLACE)
                .setPollInterval(Duration.newBuilder().setNanos(1_000_000)).build())
                .forEachRemaining(events::add);

        assertThat(events).hasSize(2);
        assertThat(events.get(0).getProgress().getPercentageComplete()).isEqualTo(40.5);
        assertThat(events.get(1).getItem().getId()).isEqualTo("copy1");
        assertThat(events.get(1).getItem().getDriveId()).isEqualTo("d2");
        FakeGraphServer.RecordedRequest copy = fake.only("/drives/d1/items/f1/copy");
        assertThat(copy.query()).isEqualTo("@microsoft.graph.conflictBehavior=replace");
        assertThat(body(copy).path("parentReference").path("driveId").asText()).isEqualTo("d2");
        assertThat(body(copy).path("parentReference").path("id").asText()).isEqualTo("p9");
        assertThat(body(copy).path("name").asText()).isEqualTo("b.txt");
        assertThat(fake.requestsTo("/monitor/1")).hasSize(2)
                .allMatch(r -> r.authorization() == null);
    }

    @Test
    void copyDriveItemReportsAFailedCopy() {
        fake.stub("POST", "/drives/d1/items/f1/copy", FakeGraphServer.empty(202,
                Map.of("Location", fake.baseUrl() + "/monitor/2")));
        fake.stub("/monitor/2", "{\"status\":\"failed\",\"error\":{\"code\":\"nameAlreadyExists\","
                + "\"message\":\"taken\"}}");
        assertThat(code(() -> stub.copyDriveItem(CopyDriveItemRequest.newBuilder()
                .setItem(item("f1")).build()).forEachRemaining(event -> { })))
                .isEqualTo(Status.Code.ABORTED);
    }

    // ------------------------------------------------------------ list items

    private static FieldValue.Builder field(String name) {
        return FieldValue.newBuilder().setName(name);
    }

    @Test
    void updateListItemFieldsWritesEveryValueType() throws Exception {
        fake.stub(LIST_ROUTE + "/items/7/fields", "{\"@odata.etag\":\"\\\"v2\\\"\","
                + "\"Title\":\"Hello\",\"Score\":4.5}");
        fake.stub("/sites/" + SITE + "/lists/User Information List/items",
                "{\"value\":[{\"id\":\"12\"}]}");
        fake.stub(LIST_ROUTE + "/columns", "{\"value\":["
                + "{\"name\":\"Topic\",\"displayName\":\"Topic\"},"
                + "{\"name\":\"a1b2\",\"displayName\":\"Topic_0\"}]}");

        UpdateListItemFieldsResponse response = stub.updateListItemFields(
                UpdateListItemFieldsRequest.newBuilder().setExpectedEtag("\"v1\"")
                        .setPatch(ListItemPatch.newBuilder()
                                .setTarget(ListItemTarget.newBuilder().setListItem(
                                        ListItemRef.newBuilder().setSiteId(SITE).setListId("l1")
                                                .setItemId("7")))
                                .addFields(field("Title").setText("Hello"))
                                .addFields(field("Body").setNote("long"))
                                .addFields(field("Score").setNumber(4.5))
                                .addFields(field("Done").setBoolean(true))
                                .addFields(field("Due").setDateTime(
                                        Timestamp.newBuilder().setSeconds(1_700_000_000)))
                                .addFields(field("Kind").setChoice("A"))
                                .addFields(field("Tags").setMultiChoice(
                                        MultiChoiceValue.newBuilder().addChoices("x").addChoices("y")))
                                .addFields(field("Topic").setTerm(TermValue.newBuilder()
                                        .setTermId("g1").setLabel("Cats")))
                                .addFields(field("Topics").setTermNoteColumn("TopicsNote")
                                        .setMultiTerm(MultiTermValue.newBuilder()
                                                .addTerms(TermValue.newBuilder().setTermId("g1")
                                                        .setLabel("Cats").setWssId(3))
                                                .addTerms(TermValue.newBuilder().setTermId("g2")
                                                        .setLabel("Dogs"))))
                                .addFields(field("Owner").setPerson(
                                        PersonValue.newBuilder().setEmail("o'neil@contoso.com")))
                                .addFields(field("Reviewers").setMultiPerson(
                                        MultiPersonValue.newBuilder()
                                                .addPeople(PersonValue.newBuilder().setLookupId(5))
                                                .addPeople(PersonValue.newBuilder().setLookupId(6))))
                                .addFields(field("Link").setHyperlink(
                                        HyperlinkValue.newBuilder().setUrl("https://example.com")))
                                .addFields(field("Parent").setLookup(9))
                                .addFields(field("Related").setMultiLookup(
                                        MultiLookupValue.newBuilder().addLookupIds(1).addLookupIds(2)))
                                .addFields(field("Old").setClear(ClearValue.getDefaultInstance())))
                        .build());

        FakeGraphServer.RecordedRequest patch = fake.only(LIST_ROUTE + "/items/7/fields");
        assertThat(patch.method()).isEqualTo("PATCH");
        assertThat(patch.ifMatch()).isEqualTo("\"v1\"");
        assertThat(body(patch)).isEqualTo(JSON.readTree("""
                {"Title":"Hello","Body":"long","Score":4.5,"Done":true,
                 "Due":"2023-11-14T22:13:20Z","Kind":"A",
                 "Tags@odata.type":"Collection(Edm.String)","Tags":["x","y"],
                 "a1b2":"-1;#Cats|g1",
                 "TopicsNote":"3;#Cats|g1;#-1;#Dogs|g2",
                 "OwnerLookupId":12,
                 "ReviewersLookupId@odata.type":"Collection(Edm.Int32)","ReviewersLookupId":[5,6],
                 "Link":{"Url":"https://example.com","Description":"https://example.com"},
                 "ParentLookupId":9,
                 "RelatedLookupId@odata.type":"Collection(Edm.Int32)","RelatedLookupId":[1,2],
                 "Old":null}
                """));
        assertThat(fake.only("/sites/" + SITE + "/lists/User Information List/items").query())
                .contains("fields%2FEMail%20eq%20%27o%27%27neil%40contoso.com%27");
        assertThat(response.getEtag()).isEqualTo("\"v2\"");
        assertThat(response.getFieldsList()).extracting(c -> c.getName())
                .containsExactly("Title", "Score");
        assertThat(response.getOutcomesList()).hasSize(15)
                .allMatch(o -> o.getState() == FieldWriteState.FIELD_WRITE_STATE_WRITTEN);
        assertThat(response.getOutcomesList()).extracting(FieldOutcome::getGraphField)
                .contains("a1b2", "TopicsNote", "OwnerLookupId", "RelatedLookupId", "Title");
    }

    @Test
    void updateListItemFieldsByDriveItemAndDryRun() throws Exception {
        fake.stub("/drives/d1/items/f1/listItem/fields", "{\"Title\":\"T\"}");
        ListItemPatch patch = ListItemPatch.newBuilder()
                .setTarget(ListItemTarget.newBuilder().setDriveItem(item("f1")))
                .addFields(field("Title").setText("T")).build();

        UpdateListItemFieldsResponse dry = stub.updateListItemFields(
                UpdateListItemFieldsRequest.newBuilder().setPatch(patch).setDryRun(true).build());
        assertThat(dry.getOutcomes(0).getState())
                .isEqualTo(FieldWriteState.FIELD_WRITE_STATE_SKIPPED);
        assertThat(fake.requests()).isEmpty();

        UpdateListItemFieldsResponse written = stub.updateListItemFields(
                UpdateListItemFieldsRequest.newBuilder().setPatch(patch).build());
        assertThat(written.getFields(0).getStringValue()).isEqualTo("T");
        assertThat(fake.only("/drives/d1/items/f1/listItem/fields").method()).isEqualTo("PATCH");

        assertThat(code(() -> stub.updateListItemFields(UpdateListItemFieldsRequest.newBuilder()
                .setPatch(patch.toBuilder().addFields(field("Title").setText("again"))).build())))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(code(() -> stub.updateListItemFields(UpdateListItemFieldsRequest.newBuilder()
                .setPatch(patch.toBuilder().clearFields().addFields(field("Empty"))).build())))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void createAndDeleteListItem() throws Exception {
        fake.stub("POST", LIST_ROUTE + "/items", FakeGraphServer.json(201,
                "{\"id\":\"8\",\"eTag\":\"\\\"n1\\\"\",\"fields\":{\"Title\":\"New\"}}", Map.of()));
        fake.stub("DELETE", LIST_ROUTE + "/items/8", FakeGraphServer.empty(204, Map.of()));
        ListRef list = ListRef.newBuilder().setSiteId(SITE).setListId("l1").build();

        var created = stub.createListItem(CreateListItemRequest.newBuilder().setList(list)
                .setContentTypeId("0x01").addFields(field("Title").setText("New")).build()).getItem();
        assertThat(created.getId()).isEqualTo("8");
        assertThat(created.getListId()).isEqualTo("l1");
        assertThat(created.getFields(0).getStringValue()).isEqualTo("New");
        JsonNode body = body(fake.only(LIST_ROUTE + "/items"));
        assertThat(body.path("fields").path("Title").asText()).isEqualTo("New");
        assertThat(body.path("contentType").path("id").asText()).isEqualTo("0x01");

        ListItemRef ref = ListItemRef.newBuilder().setSiteId(SITE).setListId("l1").setItemId("8")
                .build();
        assertThat(stub.deleteListItem(DeleteListItemRequest.newBuilder().setItem(ref)
                .setExpectedEtag("\"n1\"").build()).getItem()).isEqualTo(ref);
        assertThat(fake.only(LIST_ROUTE + "/items/8").method()).isEqualTo("DELETE");
        assertThat(fake.only(LIST_ROUTE + "/items/8").ifMatch()).isEqualTo("\"n1\"");
    }

    // ------------------------------------------------------------ provisioning

    @Test
    void ensureColumnsCreatesOnlyWhatIsMissing() throws Exception {
        fake.stub("GET", LIST_ROUTE + "/columns", FakeGraphServer.Stub.json("{\"value\":["
                + "{\"id\":\"c1\",\"name\":\"Summary\",\"text\":{\"allowMultipleLines\":true}},"
                + "{\"id\":\"c2\",\"name\":\"Kind\",\"text\":{\"allowMultipleLines\":false}}]}"));
        fake.stubOnce(LIST_ROUTE + "/columns", FakeGraphServer.Stub.json("{\"value\":["
                + "{\"id\":\"c1\",\"name\":\"Summary\",\"text\":{\"allowMultipleLines\":true}},"
                + "{\"id\":\"c2\",\"name\":\"Kind\",\"text\":{\"allowMultipleLines\":false}}]}"));
        fake.stubOnce(LIST_ROUTE + "/columns", FakeGraphServer.json(201,
                "{\"id\":\"c3\",\"name\":\"Tags\",\"choice\":{\"displayAs\":\"checkBoxes\","
                        + "\"choices\":[\"x\"]}}", Map.of()));
        fake.stubOnce(LIST_ROUTE + "/columns",
                FakeGraphServer.error(403, "accessDenied", Map.of()));
        EnsureColumnsRequest request = EnsureColumnsRequest.newBuilder()
                .setScope(ColumnScope.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(SITE).setListId("l1")))
                .addColumns(ColumnSpec.newBuilder().setName("Summary")
                        .setType(ColumnType.COLUMN_TYPE_NOTE))
                .addColumns(ColumnSpec.newBuilder().setName("Kind")
                        .setType(ColumnType.COLUMN_TYPE_CHOICE).addChoices("A"))
                .addColumns(ColumnSpec.newBuilder().setName("Tags").setDisplayName("Tag list")
                        .setType(ColumnType.COLUMN_TYPE_MULTICHOICE).addChoices("x")
                        .setAllowTextEntry(true).setIndexed(true))
                .addColumns(ColumnSpec.newBuilder().setName("Owner")
                        .setType(ColumnType.COLUMN_TYPE_USER).setMultiValued(true))
                .build();

        List<ColumnOutcome> outcomes = stub.ensureColumns(request).getColumnsList();

        assertThat(outcomes).extracting(ColumnOutcome::getState).containsExactly(
                EnsureState.ENSURE_STATE_EXISTS, EnsureState.ENSURE_STATE_DRIFT,
                EnsureState.ENSURE_STATE_CREATED, EnsureState.ENSURE_STATE_FAILED);
        assertThat(outcomes.get(1).getDriftList()).containsExactly("type");
        assertThat(outcomes.get(2).getActual().getId()).isEqualTo("c3");
        assertThat(outcomes.get(3).getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_PERMISSION_DENIED);
        List<FakeGraphServer.RecordedRequest> posts = fake.requestsTo(LIST_ROUTE + "/columns")
                .stream().filter(r -> r.method().equals("POST")).toList();
        assertThat(posts).hasSize(2);
        assertThat(body(posts.get(0))).isEqualTo(JSON.readTree("""
                {"name":"Tags","displayName":"Tag list","indexed":true,
                 "choice":{"allowTextEntry":true,"displayAs":"checkBoxes","choices":["x"]}}
                """));
        assertThat(body(posts.get(1)).path("personOrGroup").path("allowMultipleSelection")
                .asBoolean()).isTrue();

        int before = fake.requests().size();
        assertThat(stub.ensureColumns(request.toBuilder().setDryRun(true).build())
                .getColumnsList()).extracting(ColumnOutcome::getState).containsExactly(
                EnsureState.ENSURE_STATE_EXISTS, EnsureState.ENSURE_STATE_DRIFT,
                EnsureState.ENSURE_STATE_WOULD_CREATE, EnsureState.ENSURE_STATE_WOULD_CREATE);
        assertThat(fake.requests().subList(before, fake.requests().size()))
                .allMatch(r -> r.method().equals("GET"));

        assertThat(code(() -> stub.ensureColumns(request.toBuilder().addColumns(
                ColumnSpec.newBuilder().setName("Calc").setType(ColumnType.COLUMN_TYPE_CALCULATED))
                .build()))).isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void ensureListCreatesALibraryWhenMissing() throws Exception {
        String lists = "/sites/" + SITE + "/lists";
        EnsureListRequest request = EnsureListRequest.newBuilder().setSiteId(SITE)
                .setDisplayName("Team Docs").setDescription("docs")
                .setTemplate(ListTemplate.LIST_TEMPLATE_DOCUMENT_LIBRARY)
                .addColumns(ColumnSpec.newBuilder().setName("Summary")
                        .setType(ColumnType.COLUMN_TYPE_NOTE))
                .build();

        assertThat(stub.ensureList(request.toBuilder().setDryRun(true).build()).getState())
                .isEqualTo(EnsureState.ENSURE_STATE_WOULD_CREATE);
        assertThat(fake.requests()).extracting(FakeGraphServer.RecordedRequest::method)
                .containsExactly("GET");

        fake.stub("POST", lists, FakeGraphServer.json(201, "{\"id\":\"l9\",\"displayName\":"
                + "\"Team Docs\",\"list\":{\"template\":\"documentLibrary\"}}", Map.of()));
        EnsureListResponse created = stub.ensureList(request);
        assertThat(created.getState()).isEqualTo(EnsureState.ENSURE_STATE_CREATED);
        assertThat(created.getList().getId()).isEqualTo("l9");
        FakeGraphServer.RecordedRequest post = fake.requestsTo(lists).stream()
                .filter(r -> r.method().equals("POST")).findFirst().orElseThrow();
        assertThat(body(post)).isEqualTo(JSON.readTree("""
                {"displayName":"Team Docs","description":"docs",
                 "list":{"template":"documentLibrary"},
                 "columns":[{"name":"Summary","displayName":"Summary",
                             "text":{"allowMultipleLines":true}}]}
                """));

        fake.stub(lists + "/Team Docs", "{\"id\":\"l9\",\"displayName\":\"Team Docs\","
                + "\"list\":{\"template\":\"genericList\"}}");
        assertThat(stub.ensureList(request).getState()).isEqualTo(EnsureState.ENSURE_STATE_DRIFT);
        assertThat(stub.ensureList(request.toBuilder()
                .setTemplate(ListTemplate.LIST_TEMPLATE_GENERIC_LIST).build()).getState())
                .isEqualTo(EnsureState.ENSURE_STATE_EXISTS);
        assertThat(fake.requestsTo(lists).stream().filter(r -> r.method().equals("POST")))
                .hasSize(1);
    }

    // ------------------------------------------------------------ bulk

    private static MutateRequest deleteOp(String opId, String itemId, String... dependsOn) {
        return MutateRequest.newBuilder().setOpId(opId).addAllDependsOn(List.of(dependsOn))
                .setDeleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(item(itemId))
                        .setExpectedEtag("\"" + itemId + "\""))
                .build();
    }

    private List<MutateResponse> mutate(List<MutateRequest> ops) throws InterruptedException {
        Collector<MutateResponse> responses = new Collector<>();
        StreamObserver<MutateRequest> requests = async.mutate(responses);
        ops.forEach(requests::onNext);
        requests.onCompleted();
        return responses.await();
    }

    @Test
    void mutateBatchesOpsAndAnswersEachWithItsTypedResult() throws Exception {
        fake.enableBatch();
        fake.stub("POST", "/drives/d1/root/children",
                FakeGraphServer.json(201, MicrosoftFixtures.folderJson("n1", "New", "d1"), Map.of()));
        fake.stub("PATCH", "/drives/d1/items/f1",
                FakeGraphServer.Stub.json(MicrosoftFixtures.fileJson("f1", "r.txt", "d1")));
        fake.stub("DELETE", "/drives/d1/items/gone", FakeGraphServer.empty(204, Map.of()));
        fake.stub("DELETE", "/drives/d1/items/stale",
                FakeGraphServer.error(412, "preconditionFailed", Map.of()));
        fake.stub("DELETE", "/drives/d1/items/busy",
                FakeGraphServer.error(429, "tooManyRequests", Map.of("Retry-After", "9")));
        fake.stub("PATCH", LIST_ROUTE + "/items/7/fields",
                FakeGraphServer.Stub.json("{\"Title\":\"T\"}"));

        List<MutateResponse> responses = mutate(List.of(
                MutateRequest.newBuilder().setOpId("mk").setCreateFolder(
                        CreateFolderRequest.newBuilder().setDriveId("d1").setName("New")).build(),
                MutateRequest.newBuilder().setOpId("mv").addDependsOn("mk").setUpdateDriveItem(
                        UpdateDriveItemRequest.newBuilder().setItem(item("f1")).setName("r.txt"))
                        .build(),
                deleteOp("rm", "gone"),
                deleteOp("stale", "stale"),
                deleteOp("busy", "busy"),
                MutateRequest.newBuilder().setOpId("cols").setUpdateListItemFields(
                        UpdateListItemFieldsRequest.newBuilder().setPatch(ListItemPatch.newBuilder()
                                .setTarget(ListItemTarget.newBuilder().setListItem(
                                        ListItemRef.newBuilder().setSiteId(SITE).setListId("l1")
                                                .setItemId("7")))
                                .addFields(field("Title").setText("T")))).build(),
                MutateRequest.newBuilder().setOpId("bad").setCreateFolder(
                        CreateFolderRequest.newBuilder().setDriveId("d1").setName("a/b")).build(),
                MutateRequest.newBuilder().setOpId("empty").build()));

        Map<String, MutateResponse> byId = new java.util.HashMap<>();
        responses.forEach(r -> assertThat(byId.put(r.getOpId(), r)).isNull());
        assertThat(byId).hasSize(8);
        assertThat(byId.get("mk").getCreateFolder().getItem().getId()).isEqualTo("n1");
        assertThat(byId.get("mv").getUpdateDriveItem().getItem().getName()).isEqualTo("r.txt");
        assertThat(byId.get("rm").getDeleteDriveItem().getItem()).isEqualTo(item("gone"));
        assertThat(byId.get("stale").getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_ABORTED);
        assertThat(byId.get("busy").getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_RESOURCE_EXHAUSTED);
        assertThat(byId.get("busy").getError().getRetryDelay().getSeconds()).isEqualTo(9);
        assertThat(byId.get("cols").getUpdateListItemFields().getFields(0).getStringValue())
                .isEqualTo("T");
        assertThat(byId.get("bad").getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_INVALID_ARGUMENT);
        assertThat(byId.get("empty").getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_INVALID_ARGUMENT);

        List<FakeGraphServer.RecordedRequest> batches = fake.requestsTo("/$batch");
        assertThat(batches).isNotEmpty().allMatch(r -> r.method().equals("POST"));
        List<JsonNode> sent = new ArrayList<>();
        for (FakeGraphServer.RecordedRequest batch : batches) {
            body(batch).path("requests").forEach(sent::add);
        }
        assertThat(sent).hasSize(6);
        JsonNode stale = sent.stream().filter(r -> r.path("url").asText().endsWith("/stale"))
                .findFirst().orElseThrow();
        assertThat(stale.path("method").asText()).isEqualTo("DELETE");
        assertThat(stale.path("headers").path("If-Match").asText()).isEqualTo("\"stale\"");
        if (batches.size() == 1) {
            assertThat(sent.get(1).path("dependsOn").path(0).asText())
                    .isEqualTo(sent.get(0).path("id").asText());
            assertThat(sent.get(1).path("headers").path("Content-Type").asText())
                    .isEqualTo("application/json");
        }
    }

    @Test
    void mutateSendsAtMostTwentyOpsPerBatchAndFailsDependentsOfAFailedOp() throws Exception {
        fake.enableBatch();
        List<MutateRequest> ops = new ArrayList<>();
        for (int i = 0; i < 45; i++) {
            fake.stub("DELETE", "/drives/d1/items/i" + i, FakeGraphServer.empty(204, Map.of()));
            ops.add(deleteOp("op" + i, "i" + i));
        }
        fake.stub("DELETE", "/drives/d1/items/i0",
                FakeGraphServer.error(404, "itemNotFound", Map.of()));
        Collector<MutateResponse> responses = new Collector<>();
        StreamObserver<MutateRequest> requests = async.mutate(responses);
        ops.forEach(requests::onNext);
        for (int i = 0; i < 45; i++) {
            responses.next();
        }
        requests.onNext(deleteOp("after", "i1", "op0"));
        requests.onNext(deleteOp("fine", "i2", "op1"));
        requests.onCompleted();

        List<MutateResponse> all = responses.await();
        assertThat(all).hasSize(47);
        assertThat(all.stream().filter(MutateResponse::hasError).map(MutateResponse::getOpId))
                .containsExactlyInAnyOrder("op0", "after");
        MutateResponse after = all.stream().filter(r -> r.getOpId().equals("after")).findFirst()
                .orElseThrow();
        assertThat(after.getError().getCode())
                .isEqualTo(WriteErrorCode.WRITE_ERROR_CODE_FAILED_PRECONDITION);
        List<FakeGraphServer.RecordedRequest> batches = fake.requestsTo("/$batch");
        assertThat(batches.size()).isGreaterThanOrEqualTo(4);
        int total = 0;
        for (FakeGraphServer.RecordedRequest batch : batches) {
            int size = body(batch).path("requests").size();
            assertThat(size).isBetween(1, 20);
            total += size;
        }
        assertThat(total).isEqualTo(46);
    }
}
