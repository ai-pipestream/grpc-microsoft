package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ColumnOutcome;
import ai.pipestream.microsoft.v1.ColumnScope;
import ai.pipestream.microsoft.v1.ColumnSpec;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CopyDriveItemRequest;
import ai.pipestream.microsoft.v1.CopyDriveItemResponse;
import ai.pipestream.microsoft.v1.CreateFolderRequest;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.DateTimeColumnFormat;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeleteListItemRequest;
import ai.pipestream.microsoft.v1.DownloadContentRequest;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.DriveItemVersion;
import ai.pipestream.microsoft.v1.EnsureColumnsRequest;
import ai.pipestream.microsoft.v1.EnsureListRequest;
import ai.pipestream.microsoft.v1.EnsureListResponse;
import ai.pipestream.microsoft.v1.EnsureState;
import ai.pipestream.microsoft.v1.FieldValue;
import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.GetListItemRequest;
import ai.pipestream.microsoft.v1.GetListRequest;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.ListColumnsRequest;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.ListItemPatch;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemTarget;
import ai.pipestream.microsoft.v1.ListItemVersionsRequest;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListTemplate;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MutateRequest;
import ai.pipestream.microsoft.v1.MutateResponse;
import ai.pipestream.microsoft.v1.RestoreDriveItemVersionRequest;
import ai.pipestream.microsoft.v1.UpdateDriveItemRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsRequest;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import io.grpc.ManagedChannel;
import io.grpc.MethodDescriptor;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.time.Duration;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Calls every write rpc of {@code MicrosoftService} against the real tenant and
 * writes {@code build/live-write.md}.
 *
 * <p>Runs only when {@code MICROSOFT_TENANT_ID}, {@code MICROSOFT_CLIENT_ID},
 * {@code MICROSOFT_CLIENT_SECRET} and {@code MICROSOFT_SITE_ID} are set and
 * {@code MICROSOFT_LIVE_WRITE=1}. Everything it writes sits in one new folder
 * at the drive root and one new list, both removed at the end. Run with
 * {@code ./gradlew :grpc-microsoft-service:liveWriteTest}.</p>
 */
class MicrosoftLiveWriteIT {

    /** Environment switch that allows writing to the tenant. */
    static final String ENV_LIVE_WRITE = "MICROSOFT_LIVE_WRITE";

    private static final String TAG = "[MicrosoftLiveWriteIT] ";
    private static final int UNIT = GraphFiles.FRAGMENT_UNIT_BYTES;
    private static final long CALL_DEADLINE_SECONDS = 120;
    private static final long UPLOAD_DEADLINE_SECONDS = 600;

    /** The rpcs this test must exercise. */
    private static final Set<String> WRITE_RPCS = Set.of("UploadItem", "CreateFolder",
            "UpdateDriveItem", "DeleteDriveItem", "RestoreDriveItemVersion", "CopyDriveItem",
            "CreateListItem", "UpdateListItemFields", "DeleteListItem", "EnsureColumns",
            "EnsureList", "Mutate");

    private record Row(String rpc, String label, String outcome, String detail) {
    }

    /** One step of the run; answers the detail for its PASS row. */
    @FunctionalInterface
    private interface Step {
        String run() throws Exception;
    }

    /** Thrown when a step cannot run because an earlier one left nothing to work on. */
    private static final class Skip extends RuntimeException {
        private static final long serialVersionUID = 1L;

        Skip(String message) {
            super(message);
        }
    }

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

        /** The next response; fails when the call ends or stalls first. */
        T next() throws Exception {
            long deadline = System.nanoTime() + TimeUnit.SECONDS.toNanos(UPLOAD_DEADLINE_SECONDS);
            while (System.nanoTime() < deadline) {
                T value = messages.poll(200, TimeUnit.MILLISECONDS);
                if (value != null) {
                    return value;
                }
                if (done.getCount() == 0 && messages.isEmpty()) {
                    break;
                }
            }
            if (error.get() != null) {
                throw Status.fromThrowable(error.get()).asRuntimeException();
            }
            throw new IllegalStateException("the stream ended without the expected response");
        }

        /** Every response once the call has completed; rethrows its error. */
        List<T> await() throws Exception {
            if (!done.await(UPLOAD_DEADLINE_SECONDS, TimeUnit.SECONDS)) {
                throw new IllegalStateException("the call did not end in time");
            }
            if (error.get() != null) {
                throw Status.fromThrowable(error.get()).asRuntimeException();
            }
            synchronized (this) {
                return List.copyOf(all);
            }
        }

        void awaitEnd() throws InterruptedException {
            done.await(CALL_DEADLINE_SECONDS, TimeUnit.SECONDS);
        }
    }

    private final List<Row> rows = new ArrayList<>();
    private final List<String> cleanup = new ArrayList<>();
    private final Set<String> secrets = ConcurrentHashMap.newKeySet();
    private final SecureRandom random = new SecureRandom();

    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub blocking;
    private MicrosoftServiceGrpc.MicrosoftServiceStub async;
    private GraphClient client;
    private String siteId;
    private String driveId;
    private String folderName;
    private String listName;

    private String folderId;
    private String listId;
    private DriveItem small;
    private byte[] smallBytes;
    private DriveItem large;
    private DriveItem copy;
    private String itemId;
    private boolean columnsReady;

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void everyWriteRpcAgainstTheLiveTenant() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID)),
                "no live Microsoft application credentials and site in the environment; skipping");
        assumeTrue("1".equals(System.getenv(ENV_LIVE_WRITE)),
                ENV_LIVE_WRITE + " is not 1; not writing to the tenant");
        Path out = Path.of(System.getProperty("liveWrite.report", "build/live-write.md"));
        List<String> failures = run(MicrosoftConnectorConfig.fromEnvironment(), out);
        assertThat(failures).as("write rpcs that failed against the tenant (see %s)", out)
                .isEmpty();
    }

    /** Runs every step, cleans up, writes and prints the report; answers the failed rows. */
    List<String> run(MicrosoftConnectorConfig config, Path out) throws Exception {
        secrets.add(config.clientSecret());
        GraphAuth auth = new GraphAuth(config.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(issue(auth));
        client = new GraphClient(config.graphBaseUrl(), () -> {
            GraphAuth.Token current = token.get();
            if (current.expired()) {
                try {
                    current = issue(auth);
                    token.set(current);
                } catch (IOException e) {
                    throw new IllegalStateException("token refresh failed", e);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    throw new IllegalStateException("token refresh interrupted", e);
                }
            }
            return current.accessToken();
        }, AuthMode.AUTH_MODE_APPLICATION);

        String suffix = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneOffset.UTC)
                .format(Instant.now()) + "-" + HexFormat.of().toHexDigits(random.nextInt());
        folderName = "pipestream-live-write-" + suffix;
        listName = "pipestream-live-write-" + suffix;
        siteId = config.siteId();

        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, new GraphFiles(client),
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES))
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        Instant started = Instant.now();
        try {
            blocking = MicrosoftServiceGrpc.newBlockingStub(channel);
            async = MicrosoftServiceGrpc.newStub(channel);
            driveId = config.driveIds().isEmpty()
                    ? client.get("/sites/" + GraphPaths.site(siteId) + "/drive?$select=id")
                            .path("id").asText("")
                    : config.driveIds().get(0);
            if (driveId.isBlank()) {
                throw new IllegalStateException("the site has no default document library");
            }
            try {
                exercise();
            } finally {
                cleanUp();
            }
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            server.awaitTermination(10, TimeUnit.SECONDS);
            executor.shutdownNow();
        }

        for (MethodDescriptor<?, ?> method : MicrosoftServiceGrpc.getServiceDescriptor().getMethods()) {
            String rpc = method.getBareMethodName();
            if (WRITE_RPCS.contains(rpc) && rows.stream().noneMatch(row -> row.rpc().equals(rpc))) {
                rows.add(new Row(rpc, rpc, "FAIL", "not exercised by this test"));
            }
        }
        String report = report(URI.create(config.graphBaseUrl()).getHost(), started);
        Files.createDirectories(out.toAbsolutePath().getParent());
        Files.writeString(out, report, StandardCharsets.UTF_8);
        System.out.println(report);
        System.out.println(TAG + "report written to " + out.toAbsolutePath());

        List<String> failures = new ArrayList<>();
        rows.stream().filter(row -> row.outcome().equals("FAIL"))
                .forEach(row -> failures.add(scrub(row.label() + ": " + row.detail())));
        return failures;
    }

    private GraphAuth.Token issue(GraphAuth auth) throws IOException, InterruptedException {
        GraphAuth.Token issued = auth.clientCredentials();
        if (notBlank(issued.accessToken())) {
            secrets.add(issued.accessToken());
        }
        if (notBlank(issued.refreshToken())) {
            secrets.add(issued.refreshToken());
        }
        return issued;
    }

    // ------------------------------------------------------------ the run

    private void exercise() {
        step("CreateFolder", "CreateFolder", this::createFolder);
        step("UploadItem", "UploadItem [small, one PUT]", this::uploadSmall);
        step("UploadItem", "UploadItem [large, upload session]", this::uploadLarge);
        step("UploadItem", "UploadItem [resumed session]", this::uploadResumed);
        step("UpdateDriveItem", "UpdateDriveItem [rename, description]", this::updateDriveItem);
        step("UpdateDriveItem", "UpdateDriveItem [stale etag]", this::updateWithStaleEtag);
        step("CopyDriveItem", "CopyDriveItem", this::copyDriveItem);
        step("RestoreDriveItemVersion", "RestoreDriveItemVersion", this::restoreVersion);
        step("EnsureList", "EnsureList [create]", this::ensureListCreates);
        step("EnsureList", "EnsureList [again]", this::ensureListAgain);
        step("EnsureColumns", "EnsureColumns [dry run]", () -> ensureColumns(true));
        step("EnsureColumns", "EnsureColumns [create]", () -> ensureColumns(false));
        step("CreateListItem", "CreateListItem", this::createListItem);
        step("UpdateListItemFields", "UpdateListItemFields [every column type]",
                this::updateListItemFields);
        step("UpdateListItemFields", "UpdateListItemFields [drive item]",
                this::updateDriveItemFields);
        step("Mutate", "Mutate", this::mutate);
        step("DeleteListItem", "DeleteListItem", this::deleteListItem);
        step("DeleteDriveItem", "DeleteDriveItem [recycle bin]", this::recycleDriveItem);
        step("DeleteDriveItem", "DeleteDriveItem [permanent]", this::deleteForGood);
    }

    private void step(String rpc, String label, Step body) {
        String outcome;
        String detail;
        try {
            detail = body.run();
            outcome = "PASS";
        } catch (Skip e) {
            outcome = "SKIPPED";
            detail = e.getMessage();
        } catch (StatusRuntimeException e) {
            outcome = "FAIL";
            Status.Code code = e.getStatus().getCode();
            detail = (code == Status.Code.PERMISSION_DENIED || code == Status.Code.UNAUTHENTICATED
                    ? "app lacks write permission: " : "") + code + " "
                    + e.getStatus().getDescription();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            outcome = "FAIL";
            detail = "interrupted";
        } catch (Exception | AssertionError e) {
            outcome = "FAIL";
            detail = e.getClass().getSimpleName() + ": " + e.getMessage();
        }
        rows.add(new Row(rpc, label, outcome, detail));
        System.out.println(TAG + scrub(label + ": " + outcome + " " + detail));
    }

    private static <T> T need(T value, String what) {
        if (value == null) {
            throw new Skip("needs " + what + ", which an earlier step did not produce");
        }
        return value;
    }

    private static void check(boolean condition, String message) {
        if (!condition) {
            throw new AssertionError(message);
        }
    }

    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub() {
        return blocking.withDeadlineAfter(CALL_DEADLINE_SECONDS, TimeUnit.SECONDS);
    }

    // ------------------------------------------------------------ drive items

    private String createFolder() {
        DriveItem folder = stub().createFolder(CreateFolderRequest.newBuilder().setDriveId(driveId)
                .setName(folderName).setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_FAIL)
                .build()).getItem();
        folderId = folder.getId();
        DriveItem read = stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId)
                .setPath("/" + folderName).build()).getItem();
        check(read.getId().equals(folderId) && read.getFolder(), "GetItem by path did not find it");
        return "created /" + folderName;
    }

    private String uploadSmall() throws Exception {
        smallBytes = ("live write " + folderName + "\n").getBytes(StandardCharsets.UTF_8);
        small = upload("small.txt", smallBytes, 0, ConflictBehavior.CONFLICT_BEHAVIOR_FAIL);
        check(Arrays.equals(download(small.getId()), smallBytes), "downloaded bytes differ");
        return smallBytes.length + " bytes, read back equal";
    }

    private String uploadLarge() throws Exception {
        byte[] content = randomBytes(GraphFiles.SIMPLE_UPLOAD_LIMIT + 2 * UNIT + 123);
        Collector<UploadItemResponse> responses = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(responses);
        requests.onNext(frame(header("large.bin", content.length, 5 * UNIT)));
        sendChunks(requests, content, 0, content.length);
        requests.onNext(end(sha256(content)));
        requests.onCompleted();
        List<UploadItemResponse> all = responses.await();
        long progress = all.stream().filter(UploadItemResponse::hasProgress).count();
        large = all.get(all.size() - 1).getItem();
        check(progress >= 3, "expected progress per fragment, saw " + progress);
        check(sha256(download(large.getId())).equals(sha256(content)), "downloaded sha256 differs");
        return content.length + " bytes in " + (progress - 1) + " fragments, sha256 read back equal";
    }

    private String uploadResumed() throws Exception {
        byte[] content = randomBytes(14 * UNIT + 777);
        UploadItemHeader.Builder header = header("resumed.bin", content.length, 5 * UNIT);
        Collector<UploadItemResponse> first = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(first);
        requests.onNext(frame(header));
        sendChunks(requests, content, 0, 11 * UNIT);
        String token = "";
        long committed = 0;
        while (committed < 10L * UNIT) {
            UploadItemResponse progress = first.next();
            token = progress.getProgress().getResumeToken();
            committed = progress.getProgress().getCommittedBytes();
        }
        requests.onError(Status.CANCELLED.withDescription("interrupted on purpose")
                .asRuntimeException());
        first.awaitEnd();

        Collector<UploadItemResponse> second = new Collector<>();
        requests = async.uploadItem(second);
        requests.onNext(frame(header.setResumeToken(token).setResumeOffset(committed)));
        sendChunks(requests, content, (int) committed, content.length);
        requests.onNext(end(""));
        requests.onCompleted();
        List<UploadItemResponse> all = second.await();
        DriveItem item = all.get(all.size() - 1).getItem();
        check(sha256(download(item.getId())).equals(sha256(content)), "downloaded sha256 differs");
        return "interrupted at " + committed + " of " + content.length
                + " bytes, resumed, sha256 read back equal";
    }

    private String updateDriveItem() {
        need(small, "the small upload");
        String description = "renamed by the live write test";
        DriveItem updated = stub().updateDriveItem(UpdateDriveItemRequest.newBuilder()
                .setItem(ref(small.getId())).setExpectedEtag(small.getEtag())
                .setName("renamed.txt").setDescription(description).build()).getItem();
        DriveItem read = getItem(small.getId());
        check(read.getName().equals("renamed.txt"), "name read back as " + read.getName());
        check(read.getDescription().equals(description),
                "description read back as '" + read.getDescription() + "'");
        check(!updated.getEtag().equals(small.getEtag()), "the eTag did not change");
        return "name and description read back; eTag changed";
    }

    private String updateWithStaleEtag() {
        need(small, "the small upload");
        check(!getItem(small.getId()).getEtag().equals(small.getEtag()),
                "the item still has its first eTag, so there is no stale one to send");
        try {
            stub().updateDriveItem(UpdateDriveItemRequest.newBuilder().setItem(ref(small.getId()))
                    .setExpectedEtag(small.getEtag()).setDescription("must not land").build());
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == Status.Code.ABORTED) {
                check(!getItem(small.getId()).getDescription().equals("must not land"),
                        "the update landed despite the stale eTag");
                return "ABORTED, item unchanged";
            }
            throw e;
        }
        throw new AssertionError("the update with a stale eTag was accepted");
    }

    private String copyDriveItem() {
        need(small, "the small upload");
        need(folderId, "the test folder");
        List<CopyDriveItemResponse> events = new ArrayList<>();
        blocking.withDeadlineAfter(UPLOAD_DEADLINE_SECONDS, TimeUnit.SECONDS)
                .copyDriveItem(CopyDriveItemRequest.newBuilder().setItem(ref(small.getId()))
                        .setDestinationParentId(folderId).setName("copy.txt").build())
                .forEachRemaining(events::add);
        copy = events.get(events.size() - 1).getItem();
        check(!copy.getId().equals(small.getId()), "the copy has the source's id");
        check(getItem(copy.getId()).getName().equals("copy.txt"), "the copy has another name");
        check(Arrays.equals(download(copy.getId()), smallBytes), "the copy's bytes differ");
        return (events.size() - 1) + " progress events, copy read back equal";
    }

    private String restoreVersion() throws Exception {
        byte[] first = "version one\n".getBytes(StandardCharsets.UTF_8);
        byte[] second = "version two, longer\n".getBytes(StandardCharsets.UTF_8);
        DriveItem item = upload("versioned.txt", first, 0, ConflictBehavior.CONFLICT_BEHAVIOR_FAIL);
        upload("versioned.txt", second, 0, ConflictBehavior.CONFLICT_BEHAVIOR_REPLACE);
        List<DriveItemVersion> versions = new ArrayList<>();
        stub().listItemVersions(ListItemVersionsRequest.newBuilder().setDriveId(driveId)
                .setItemId(item.getId()).build())
                .forEachRemaining(response -> versions.add(response.getVersion()));
        check(versions.size() >= 2, "the library kept " + versions.size()
                + " version(s); versioning must be on to restore one");
        String oldest = versions.get(versions.size() - 1).getId();
        stub().restoreDriveItemVersion(RestoreDriveItemVersionRequest.newBuilder()
                .setItem(ref(item.getId())).setVersionId(oldest).build());
        check(Arrays.equals(download(item.getId()), first),
                "the content is not the first version's");
        return "restored version " + oldest + ", first content read back";
    }

    private String recycleDriveItem() {
        need(small, "the small upload");
        String etag = getItem(small.getId()).getEtag();
        stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(small.getId()))
                .setExpectedEtag(etag).build());
        check(gone(() -> getItem(small.getId())), "the item is still readable");
        return "deleted with its eTag; GetItem answers NOT_FOUND";
    }

    private String deleteForGood() {
        need(large, "the large upload");
        stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(large.getId()))
                .setPermanent(true).build());
        check(gone(() -> getItem(large.getId())), "the item is still readable");
        return "GetItem answers NOT_FOUND";
    }

    // ------------------------------------------------------------ lists

    private EnsureListRequest ensureListRequest() {
        return EnsureListRequest.newBuilder().setSiteId(siteId).setDisplayName(listName)
                .setDescription("Created by the live write test; safe to delete.")
                .setTemplate(ListTemplate.LIST_TEMPLATE_GENERIC_LIST).build();
    }

    private String ensureListCreates() {
        EnsureListResponse dry = stub().ensureList(ensureListRequest().toBuilder().setDryRun(true)
                .build());
        check(dry.getState() == EnsureState.ENSURE_STATE_WOULD_CREATE,
                "dry run answered " + dry.getState());
        EnsureListResponse created = stub().ensureList(ensureListRequest());
        check(created.getState() == EnsureState.ENSURE_STATE_CREATED,
                "answered " + created.getState());
        listId = created.getList().getId();
        check(stub().getList(GetListRequest.newBuilder().setSiteId(siteId).setListId(listId).build())
                .getList().getDisplayName().equals(listName), "GetList shows another name");
        return "dry run WOULD_CREATE, then CREATED; GetList finds it";
    }

    private String ensureListAgain() {
        need(listId, "the test list");
        EnsureListResponse again = stub().ensureList(ensureListRequest());
        check(again.getState() == EnsureState.ENSURE_STATE_EXISTS, "answered " + again.getState());
        check(again.getList().getId().equals(listId), "answered another list");
        return "EXISTS, same list";
    }

    private String ensureColumns(boolean dryRun) {
        need(listId, "the test list");
        ColumnSpec.Builder column = ColumnSpec.newBuilder();
        EnsureColumnsRequest request = EnsureColumnsRequest.newBuilder().setDryRun(dryRun)
                .setScope(ColumnScope.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId)))
                .addColumns(column.clone().setName("PsText").setType(ColumnType.COLUMN_TYPE_TEXT))
                .addColumns(column.clone().setName("PsNumber")
                        .setType(ColumnType.COLUMN_TYPE_NUMBER))
                .addColumns(column.clone().setName("PsFlag")
                        .setType(ColumnType.COLUMN_TYPE_BOOLEAN))
                .addColumns(column.clone().setName("PsChoice")
                        .setType(ColumnType.COLUMN_TYPE_CHOICE).addChoices("A").addChoices("B"))
                .addColumns(column.clone().setName("PsWhen")
                        .setType(ColumnType.COLUMN_TYPE_DATE_TIME).setDateTimeFormat(
                                DateTimeColumnFormat.DATE_TIME_COLUMN_FORMAT_DATE_TIME))
                .build();
        EnsureState wanted = dryRun ? EnsureState.ENSURE_STATE_WOULD_CREATE
                : EnsureState.ENSURE_STATE_CREATED;
        List<ColumnOutcome> outcomes = stub().ensureColumns(request).getColumnsList();
        for (ColumnOutcome outcome : outcomes) {
            if (outcome.getError().getCode() == WriteErrorCode.WRITE_ERROR_CODE_PERMISSION_DENIED
                    || outcome.getError().getCode()
                            == WriteErrorCode.WRITE_ERROR_CODE_UNAUTHENTICATED) {
                throw new AssertionError("app lacks write permission: " + outcome.getName() + " "
                        + outcome.getError().getMessage());
            }
            check(outcome.getState() == wanted, outcome.getName() + " answered "
                    + outcome.getState() + " " + outcome.getError().getMessage());
        }
        List<String> names = new ArrayList<>();
        stub().listColumns(ListColumnsRequest.newBuilder().setSiteId(siteId).setListId(listId)
                .build()).forEachRemaining(response -> names.add(response.getColumn().getName()));
        boolean present = names.containsAll(
                List.of("PsText", "PsNumber", "PsFlag", "PsChoice", "PsWhen"));
        check(present != dryRun, dryRun ? "the dry run created columns"
                : "ListColumns does not show every column");
        columnsReady = !dryRun;
        return outcomes.size() + " columns " + wanted + (dryRun ? "; none created"
                : "; ListColumns shows them");
    }

    private String createListItem() {
        need(listId, "the test list");
        ListItem created = stub().createListItem(CreateListItemRequest.newBuilder()
                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId))
                .addFields(FieldValue.newBuilder().setName("Title").setText("live write"))
                .build()).getItem();
        itemId = created.getId();
        check("live write".equals(fields(itemId).get("Title")), "Title read back differently");
        return "item " + itemId + ", Title read back";
    }

    private String updateListItemFields() {
        need(itemId, "the list item");
        if (!columnsReady) {
            throw new Skip("needs the test columns, which an earlier step did not produce");
        }
        long when = Instant.parse("2024-01-15T12:00:00Z").getEpochSecond();
        String etag = listItem(itemId).getEtag();
        stub().updateListItemFields(UpdateListItemFieldsRequest.newBuilder().setExpectedEtag(etag)
                .setPatch(ListItemPatch.newBuilder().setTarget(listItemTarget(itemId))
                        .addFields(FieldValue.newBuilder().setName("PsText").setText("hello"))
                        .addFields(FieldValue.newBuilder().setName("PsNumber").setNumber(4.5))
                        .addFields(FieldValue.newBuilder().setName("PsFlag").setBoolean(true))
                        .addFields(FieldValue.newBuilder().setName("PsChoice").setChoice("B"))
                        .addFields(FieldValue.newBuilder().setName("PsWhen").setDateTime(
                                Timestamp.newBuilder().setSeconds(when))))
                .build());
        Map<String, String> read = fields(itemId);
        check("hello".equals(read.get("PsText")), "PsText read back as " + read.get("PsText"));
        check("4.5".equals(read.get("PsNumber")), "PsNumber read back as " + read.get("PsNumber"));
        check("true".equals(read.get("PsFlag")), "PsFlag read back as " + read.get("PsFlag"));
        check("B".equals(read.get("PsChoice")), "PsChoice read back as " + read.get("PsChoice"));
        check(Long.toString(when).equals(read.get("PsWhen")),
                "PsWhen read back as " + read.get("PsWhen"));
        return "text, number, boolean, choice and date-time read back";
    }

    private String updateDriveItemFields() {
        need(small, "the small upload");
        stub().updateListItemFields(UpdateListItemFieldsRequest.newBuilder()
                .setPatch(ListItemPatch.newBuilder()
                        .setTarget(ListItemTarget.newBuilder().setDriveItem(ref(small.getId())))
                        .addFields(FieldValue.newBuilder().setName("Title").setText("live title")))
                .build());
        String title = "";
        for (ListColumn column : getItem(small.getId()).getListColumnsList()) {
            if (column.getName().equals("Title")) {
                title = column.getStringValue();
            }
        }
        check(title.equals("live title"), "Title read back as '" + title + "'");
        return "Title of a library file read back";
    }

    private String deleteListItem() {
        need(itemId, "the list item");
        String etag = listItem(itemId).getEtag();
        stub().deleteListItem(DeleteListItemRequest.newBuilder().setItem(listItemRef(itemId))
                .setExpectedEtag(etag).build());
        check(gone(() -> listItem(itemId)), "the item is still readable");
        return "deleted with its eTag; GetListItem answers NOT_FOUND";
    }

    // ------------------------------------------------------------ bulk

    private String mutate() throws Exception {
        need(folderId, "the test folder");
        need(listId, "the test list");
        need(itemId, "the list item");
        need(copy, "the copy");
        ListRef list = ListRef.newBuilder().setSiteId(siteId).setListId(listId).build();
        String doomed = stub().createListItem(CreateListItemRequest.newBuilder().setList(list)
                .addFields(FieldValue.newBuilder().setName("Title").setText("to delete")).build())
                .getItem().getId();
        List<MutateRequest> ops = List.of(
                MutateRequest.newBuilder().setOpId("folder-a").setCreateFolder(
                        CreateFolderRequest.newBuilder().setDriveId(driveId).setParentId(folderId)
                                .setName("batch-a")).build(),
                MutateRequest.newBuilder().setOpId("folder-b").addDependsOn("folder-a")
                        .setCreateFolder(CreateFolderRequest.newBuilder().setDriveId(driveId)
                                .setParentId(folderId).setName("batch-b")).build(),
                MutateRequest.newBuilder().setOpId("new-item").setCreateListItem(
                        CreateListItemRequest.newBuilder().setList(list).addFields(
                                FieldValue.newBuilder().setName("Title").setText("from batch")))
                        .build(),
                MutateRequest.newBuilder().setOpId("retitle").setUpdateListItemFields(
                        UpdateListItemFieldsRequest.newBuilder().setPatch(ListItemPatch.newBuilder()
                                .setTarget(listItemTarget(itemId)).addFields(FieldValue.newBuilder()
                                        .setName("Title").setText("retitled in batch")))).build(),
                MutateRequest.newBuilder().setOpId("drop-item").setDeleteListItem(
                        DeleteListItemRequest.newBuilder().setItem(listItemRef(doomed))).build(),
                MutateRequest.newBuilder().setOpId("describe").setUpdateDriveItem(
                        UpdateDriveItemRequest.newBuilder().setItem(ref(copy.getId()))
                                .setDescription("described in batch")).build(),
                MutateRequest.newBuilder().setOpId("drop-copy").addDependsOn("describe")
                        .setDeleteDriveItem(DeleteDriveItemRequest.newBuilder()
                                .setItem(ref(copy.getId()))).build());
        Collector<MutateResponse> responses = new Collector<>();
        StreamObserver<MutateRequest> requests = async.mutate(responses);
        ops.forEach(requests::onNext);
        requests.onCompleted();
        Map<String, MutateResponse> byId = new HashMap<>();
        responses.await().forEach(response -> byId.put(response.getOpId(), response));
        check(byId.size() == ops.size(), "answered " + byId.size() + " of " + ops.size() + " ops");
        for (MutateResponse response : byId.values()) {
            if (response.hasError()) {
                WriteErrorCode code = response.getError().getCode();
                throw new AssertionError((code == WriteErrorCode.WRITE_ERROR_CODE_PERMISSION_DENIED
                        || code == WriteErrorCode.WRITE_ERROR_CODE_UNAUTHENTICATED
                        ? "app lacks write permission: " : "") + "op " + response.getOpId() + " "
                        + code + " " + response.getError().getMessage());
            }
        }
        check(stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId)
                .setPath("/" + folderName + "/batch-b").build()).getItem().getFolder(),
                "batch-b is not a folder");
        check("from batch".equals(fields(byId.get("new-item").getCreateListItem().getItem()
                .getId()).get("Title")), "the batch-created item reads back differently");
        check("retitled in batch".equals(fields(itemId).get("Title")),
                "the batch update reads back differently");
        check(gone(() -> listItem(doomed)), "the batch-deleted list item is still readable");
        check(gone(() -> getItem(copy.getId())), "the batch-deleted copy is still readable");
        return ops.size() + " ops, two depends_on pairs, every result read back";
    }

    // ------------------------------------------------------------ cleanup

    /** Removes the test folder and the test list, and nothing else. */
    private void cleanUp() {
        try {
            String id = folderId;
            if (id == null) {
                id = stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId)
                        .setPath("/" + folderName).build()).getItem().getId();
            }
            try {
                stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(id))
                        .setPermanent(true).build());
                cleanup.add("folder /" + folderName + ": deleted");
            } catch (StatusRuntimeException permanent) {
                stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(id))
                        .build());
                cleanup.add("folder /" + folderName + ": moved to the recycle bin (permanent "
                        + "delete answered " + permanent.getStatus().getCode() + ")");
            }
        } catch (StatusRuntimeException e) {
            cleanup.add(e.getStatus().getCode() == Status.Code.NOT_FOUND && folderId == null
                    ? "folder /" + folderName + ": never created"
                    : "folder /" + folderName + ": NOT CLEANED UP, " + e.getStatus());
        }
        String lists = "/sites/" + GraphPaths.site(siteId) + "/lists/";
        try {
            String id = listId;
            if (id == null) {
                id = client.get(lists + GraphPaths.segment(listName, "list") + "?$select=id")
                        .path("id").asText("");
            }
            client.delete(lists + GraphPaths.segment(id, "list_id"));
            cleanup.add("list " + listName + ": deleted");
        } catch (GraphClient.GraphApiException e) {
            cleanup.add(e.status() == 404 && listId == null ? "list " + listName + ": never created"
                    : "list " + listName + ": NOT CLEANED UP, " + e.getMessage());
        } catch (IOException | RuntimeException e) {
            cleanup.add("list " + listName + ": NOT CLEANED UP, " + e.getMessage());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            cleanup.add("list " + listName + ": NOT CLEANED UP, interrupted");
        }
        cleanup.forEach(line -> System.out.println(TAG + "cleanup: " + scrub(line)));
    }

    // ------------------------------------------------------------ helpers

    private DriveItemRef ref(String id) {
        return DriveItemRef.newBuilder().setDriveId(driveId).setItemId(id).build();
    }

    private ListItemRef listItemRef(String id) {
        return ListItemRef.newBuilder().setSiteId(siteId).setListId(listId).setItemId(id).build();
    }

    private ListItemTarget listItemTarget(String id) {
        return ListItemTarget.newBuilder().setListItem(listItemRef(id)).build();
    }

    private DriveItem getItem(String id) {
        return stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId).setItemId(id).build())
                .getItem();
    }

    private ListItem listItem(String id) {
        return stub().getListItem(GetListItemRequest.newBuilder().setSiteId(siteId)
                .setListId(listId).setItemId(id).build()).getItem();
    }

    /** The column values of a list item, each as text. */
    private Map<String, String> fields(String id) {
        Map<String, String> values = new HashMap<>();
        for (ListColumn column : listItem(id).getFieldsList()) {
            values.put(column.getName(), switch (column.getValueCase()) {
                case STRING_VALUE -> column.getStringValue();
                case INT_VALUE -> Long.toString(column.getIntValue());
                case DOUBLE_VALUE -> Double.toString(column.getDoubleValue());
                case BOOL_VALUE -> Boolean.toString(column.getBoolValue());
                case TIMESTAMP_VALUE -> Long.toString(column.getTimestampValue().getSeconds());
                case VALUE_NOT_SET -> "";
            });
        }
        return values;
    }

    private static boolean gone(Runnable read) {
        try {
            read.run();
            return false;
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == Status.Code.NOT_FOUND) {
                return true;
            }
            throw e;
        }
    }

    private byte[] download(String id) {
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        blocking.withDeadlineAfter(UPLOAD_DEADLINE_SECONDS, TimeUnit.SECONDS)
                .downloadContent(DownloadContentRequest.newBuilder().setDriveId(driveId)
                        .setItemId(id).build())
                .forEachRemaining(response -> bytes.writeBytes(response.getChunk().toByteArray()));
        return bytes.toByteArray();
    }

    private UploadItemHeader.Builder header(String name, long size, int fragmentBytes) {
        return UploadItemHeader.newBuilder().setDriveId(driveId)
                .setParentId(need(folderId, "the test folder")).setName(name).setSize(size)
                .setFragmentBytes(fragmentBytes);
    }

    private static UploadItemRequest frame(UploadItemHeader.Builder header) {
        return UploadItemRequest.newBuilder().setHeader(header).build();
    }

    private static UploadItemRequest end(String sha256) {
        return UploadItemRequest.newBuilder().setEnd(UploadItemEnd.newBuilder().setSha256(sha256))
                .build();
    }

    private static void sendChunks(StreamObserver<UploadItemRequest> requests, byte[] content,
            int from, int to) {
        int chunk = 256 * 1024;
        for (int offset = from; offset < to; offset += chunk) {
            requests.onNext(UploadItemRequest.newBuilder()
                    .setChunk(ByteString.copyFrom(content, offset, Math.min(chunk, to - offset)))
                    .build());
        }
    }

    /** Uploads a whole file of known size and answers the stored item. */
    private DriveItem upload(String name, byte[] content, int fragmentBytes,
            ConflictBehavior conflict) throws Exception {
        Collector<UploadItemResponse> responses = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(responses);
        requests.onNext(frame(
                header(name, content.length, fragmentBytes).setConflictBehavior(conflict)));
        sendChunks(requests, content, 0, content.length);
        requests.onNext(end(sha256(content)));
        requests.onCompleted();
        List<UploadItemResponse> all = responses.await();
        return all.get(all.size() - 1).getItem();
    }

    private byte[] randomBytes(int length) {
        byte[] content = new byte[length];
        random.nextBytes(content);
        return content;
    }

    private static String sha256(byte[] content) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
    }

    // ------------------------------------------------------------ report

    private String report(String host, Instant started) {
        StringBuilder md = new StringBuilder();
        md.append("# Microsoft Graph live write run\n\n");
        md.append("- Graph: ").append(host).append(" (application token, client credentials)\n");
        md.append("- Run: ").append(started).append(" (")
                .append(Duration.between(started, Instant.now()).toSeconds()).append(" s)\n");
        md.append("- Test folder: /").append(folderName).append(" on drive ").append(driveId)
                .append("\n- Test list: ").append(listName).append("\n");
        long failed = rows.stream().filter(row -> row.outcome().equals("FAIL")).count();
        md.append("- Rows: ").append(rows.size()).append(", failed ").append(failed).append("\n\n");
        md.append("## RPCs\n\n| rpc | outcome | detail |\n|---|---|---|\n");
        for (Row row : rows) {
            md.append("| ").append(cell(row.label())).append(" | ")
                    .append(row.outcome().equals("FAIL") ? "**FAIL**" : row.outcome())
                    .append(" | ").append(cell(row.detail())).append(" |\n");
        }
        md.append("\n## Cleanup\n\n");
        cleanup.forEach(line -> md.append("- ").append(line).append("\n"));
        return scrub(md.toString());
    }

    private static String cell(String text) {
        String line = String.valueOf(text).replace('\n', ' ').replace('\r', ' ');
        return (line.length() > 400 ? line.substring(0, 400) + "..." : line).replace("|", "\\|");
    }

    private String scrub(String text) {
        String out = text;
        for (String secret : secrets) {
            if (notBlank(secret)) {
                out = out.replace(secret, "***");
            }
        }
        return out;
    }
}
