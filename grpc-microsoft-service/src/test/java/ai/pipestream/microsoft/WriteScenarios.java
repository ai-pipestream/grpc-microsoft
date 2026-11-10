package ai.pipestream.microsoft;

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
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListColumnsRequest;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.ListItemPatch;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemTarget;
import ai.pipestream.microsoft.v1.ListItemVersionsRequest;
import ai.pipestream.microsoft.v1.ListListItemsRequest;
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
import io.grpc.StatusRuntimeException;
import io.grpc.stub.StreamObserver;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;


/**
 * The write scenarios of {@code MicrosoftService}: each calls one write rpc and
 * reads the result back. They run in order and share what earlier ones created,
 * all inside one test folder and one test list. The same scenarios run against
 * the real tenant and against its recorded exchanges.
 */
final class WriteScenarios {

    private static final int UNIT = GraphFiles.FRAGMENT_UNIT_BYTES;
    private static final long CALL_DEADLINE_SECONDS = 120;
    private static final long UPLOAD_DEADLINE_SECONDS = 600;

    /**
     * {@code Mutate} batch wait both modes run with: long enough that the ops of
     * the mutate scenario always go out as one batch.
     */
    static final long MUTATE_LINGER_MILLIS = 5_000;

    /**
     * What a run works on.
     *
     * @param siteId the site of the test list
     * @param driveId the drive of the test folder
     * @param suffix end of the test folder and list names; also seeds the content
     * @param copyPollInterval time between copy progress checks; {@code null} for the default
     * @param blobs every content the scenarios produced, by sha256
     */
    record Backend(String siteId, String driveId, String suffix, Duration copyPollInterval,
            Map<String, byte[]> blobs) {
    }

    /** How a scenario ended. */
    enum Status {
        /** The rpc worked and the result read back as expected. */
        PASS,
        /** The rpc or the read-back failed. */
        FAIL,
        /** An earlier scenario left nothing to work on. */
        SKIPPED
    }

    /**
     * The result of one scenario.
     *
     * @param status how it ended
     * @param detail what was checked, or why it did not pass
     */
    record Outcome(Status status, String detail) {
    }

    /** The body of a scenario; answers the detail of its PASS. */
    @FunctionalInterface
    interface Step {
        String run() throws Exception;
    }

    /**
     * One scenario.
     *
     * @param id directory name of its recording
     * @param rpc the rpc it exercises
     * @param label name for reports
     * @param step what it does
     */
    record Scenario(String id, String rpc, String label, Step step) {
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
                throw io.grpc.Status.fromThrowable(error.get()).asRuntimeException();
            }
            throw new IllegalStateException("the stream ended without the expected response");
        }

        /** Every response once the call has completed; rethrows its error. */
        List<T> await() throws Exception {
            if (!done.await(UPLOAD_DEADLINE_SECONDS, TimeUnit.SECONDS)) {
                throw new IllegalStateException("the call did not end in time");
            }
            if (error.get() != null) {
                throw io.grpc.Status.fromThrowable(error.get()).asRuntimeException();
            }
            synchronized (this) {
                return List.copyOf(all);
            }
        }

        void awaitEnd() throws InterruptedException {
            done.await(CALL_DEADLINE_SECONDS, TimeUnit.SECONDS);
        }
    }

    private final MicrosoftServiceGrpc.MicrosoftServiceBlockingStub blocking;
    private final MicrosoftServiceGrpc.MicrosoftServiceStub async;
    private final GraphClient client;
    private final Backend backend;
    private final String siteId;
    private final String driveId;
    private final String folderName;
    private final String listName;

    private String folderId;
    private String listId;
    private DriveItem small;
    private byte[] smallBytes;
    private DriveItem large;
    private DriveItem copy;
    private String itemId;
    private boolean columnsReady;

    /**
     * Creates the scenarios of one run.
     *
     * @param blocking the service, for unary and server-streaming rpcs
     * @param async the service, for client-streaming rpcs
     * @param client the Graph client behind the service; deletes the test list
     * @param backend what the run works on
     */
    WriteScenarios(MicrosoftServiceGrpc.MicrosoftServiceBlockingStub blocking,
            MicrosoftServiceGrpc.MicrosoftServiceStub async, GraphClient client, Backend backend) {
        this.blocking = blocking;
        this.async = async;
        this.client = client;
        this.backend = backend;
        this.siteId = backend.siteId();
        this.driveId = backend.driveId();
        this.folderName = "pipestream-live-write-" + backend.suffix();
        this.listName = "pipestream-live-write-" + backend.suffix();
    }

    /** Name of the test folder at the drive root. */
    String folderName() {
        return folderName;
    }

    /** Display name of the test list. */
    String listName() {
        return listName;
    }

    /** Every scenario, in the order they must run; the last one cleans up. */
    List<Scenario> all() {
        return List.of(
                new Scenario("create-folder", "CreateFolder", "CreateFolder", this::createFolder),
                new Scenario("upload-small", "UploadItem", "UploadItem [small, one PUT]",
                        this::uploadSmall),
                new Scenario("upload-large", "UploadItem", "UploadItem [large, upload session]",
                        this::uploadLarge),
                new Scenario("upload-resumed", "UploadItem", "UploadItem [resumed session]",
                        this::uploadResumed),
                new Scenario("list-children", "ListChildren", "ListChildren [test folder]",
                        this::listChildren),
                new Scenario("update-drive-item", "UpdateDriveItem",
                        "UpdateDriveItem [rename, description]", this::updateDriveItem),
                new Scenario("update-drive-item-stale-etag", "UpdateDriveItem",
                        "UpdateDriveItem [stale etag]", this::updateWithStaleEtag),
                new Scenario("copy-drive-item", "CopyDriveItem", "CopyDriveItem",
                        this::copyDriveItem),
                new Scenario("restore-drive-item-version", "RestoreDriveItemVersion",
                        "RestoreDriveItemVersion", this::restoreVersion),
                new Scenario("ensure-list", "EnsureList", "EnsureList [create]",
                        this::ensureListCreates),
                new Scenario("ensure-list-again", "EnsureList", "EnsureList [again]",
                        this::ensureListAgain),
                new Scenario("ensure-columns-dry-run", "EnsureColumns", "EnsureColumns [dry run]",
                        () -> ensureColumns(true)),
                new Scenario("ensure-columns", "EnsureColumns", "EnsureColumns [create]",
                        () -> ensureColumns(false)),
                new Scenario("create-list-item", "CreateListItem", "CreateListItem",
                        this::createListItem),
                new Scenario("list-list-items", "ListListItems", "ListListItems [test list]",
                        this::listListItems),
                new Scenario("update-list-item-fields", "UpdateListItemFields",
                        "UpdateListItemFields [every column type]", this::updateListItemFields),
                new Scenario("update-list-item-fields-drive-item", "UpdateListItemFields",
                        "UpdateListItemFields [drive item]", this::updateDriveItemFields),
                new Scenario("mutate", "Mutate", "Mutate", this::mutate),
                new Scenario("delete-list-item", "DeleteListItem", "DeleteListItem",
                        this::deleteListItem),
                new Scenario("delete-drive-item", "DeleteDriveItem",
                        "DeleteDriveItem [recycle bin]", this::recycleDriveItem),
                new Scenario("delete-drive-item-permanent", "DeleteDriveItem",
                        "DeleteDriveItem [permanent]", this::deleteForGood),
                new Scenario("cleanup", "Cleanup", "Cleanup [test folder and list]",
                        this::cleanUp));
    }

    /** Runs one scenario; never throws. */
    Outcome run(Scenario scenario) {
        try {
            return new Outcome(Status.PASS, scenario.step().run());
        } catch (Skip e) {
            return new Outcome(Status.SKIPPED, e.getMessage());
        } catch (StatusRuntimeException e) {
            io.grpc.Status.Code code = e.getStatus().getCode();
            return new Outcome(Status.FAIL, (code == io.grpc.Status.Code.PERMISSION_DENIED
                    || code == io.grpc.Status.Code.UNAUTHENTICATED
                    ? "app lacks write permission: " : "") + code + " "
                    + e.getStatus().getDescription());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            return new Outcome(Status.FAIL, "interrupted");
        } catch (Exception | AssertionError e) {
            return new Outcome(Status.FAIL, e.getClass().getSimpleName() + ": " + e.getMessage());
        }
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
        smallBytes = remember(("live write " + folderName + "\n").getBytes(StandardCharsets.UTF_8));
        small = upload("small.txt", smallBytes, 0, ConflictBehavior.CONFLICT_BEHAVIOR_FAIL);
        check(Arrays.equals(download(small.getId()), smallBytes), "downloaded bytes differ");
        return smallBytes.length + " bytes, read back equal";
    }

    private String uploadLarge() throws Exception {
        byte[] content = content("large.bin", GraphFiles.SIMPLE_UPLOAD_LIMIT + 2 * UNIT + 123);
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
        byte[] content = content("resumed.bin", 14 * UNIT + 777);
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
        requests.onError(io.grpc.Status.CANCELLED.withDescription("interrupted on purpose")
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
            if (e.getStatus().getCode() == io.grpc.Status.Code.ABORTED) {
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
                .copyDriveItem(copyRequest())
                .forEachRemaining(events::add);
        copy = events.get(events.size() - 1).getItem();
        check(!copy.getId().equals(small.getId()), "the copy has the source's id");
        check(getItem(copy.getId()).getName().equals("copy.txt"), "the copy has another name");
        check(Arrays.equals(download(copy.getId()), smallBytes), "the copy's bytes differ");
        return (events.size() - 1) + " progress events, copy read back equal";
    }

    private CopyDriveItemRequest copyRequest() {
        CopyDriveItemRequest.Builder request = CopyDriveItemRequest.newBuilder()
                .setItem(ref(small.getId())).setDestinationParentId(folderId).setName("copy.txt");
        if (backend.copyPollInterval() != null) {
            request.setPollInterval(com.google.protobuf.Duration.newBuilder()
                    .setSeconds(backend.copyPollInterval().getSeconds())
                    .setNanos(backend.copyPollInterval().getNano()));
        }
        return request.build();
    }

    private String restoreVersion() throws Exception {
        byte[] first = remember("version one\n".getBytes(StandardCharsets.UTF_8));
        byte[] second = remember("version two, longer\n".getBytes(StandardCharsets.UTF_8));
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
            if (e.getStatus().getCode() == io.grpc.Status.Code.NOT_FOUND) {
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

    /** Content for the file {@code name}: the same bytes on every run with this suffix. */
    private byte[] content(String name, int length) {
        byte[] bytes = new byte[length];
        new Random((backend.suffix() + "/" + name).hashCode()).nextBytes(bytes);
        return remember(bytes);
    }

    private byte[] remember(byte[] bytes) {
        backend.blobs().put(Recordings.sha256(bytes), bytes);
        return bytes;
    }

    private static String sha256(byte[] content) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
    }

    // ------------------------------------------------------------ reads

    private String listChildren() {
        need(folderId, "the test folder");
        List<String> names = new ArrayList<>();
        stub().listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId)
                .setFolderId(folderId).build())
                .forEachRemaining(response -> names.add(response.getItem().getName()));
        check(names.containsAll(List.of("small.txt", "large.bin", "resumed.bin")),
                "the folder lists " + names);
        return names.size() + " children, the uploads among them";
    }

    private String listListItems() {
        need(itemId, "the list item");
        List<String> ids = new ArrayList<>();
        stub().listListItems(ListListItemsRequest.newBuilder().setSiteId(siteId).setListId(listId)
                .build()).forEachRemaining(response -> ids.add(response.getItem().getId()));
        check(ids.contains(itemId), "the list holds items " + ids);
        return ids.size() + " item(s), the created one among them";
    }

    // ------------------------------------------------------------ cleanup

    /** Removes the test folder and the test list, and nothing else. */
    private String cleanUp() throws InterruptedException {
        List<String> lines = new ArrayList<>();
        boolean clean = true;
        try {
            String id = folderId;
            if (id == null) {
                id = stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId)
                        .setPath("/" + folderName).build()).getItem().getId();
            }
            try {
                stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(id))
                        .setPermanent(true).build());
                lines.add("folder /" + folderName + " deleted");
            } catch (StatusRuntimeException permanent) {
                stub().deleteDriveItem(DeleteDriveItemRequest.newBuilder().setItem(ref(id))
                        .build());
                lines.add("folder /" + folderName + " moved to the recycle bin (permanent delete "
                        + "answered " + permanent.getStatus().getCode() + ")");
            }
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == io.grpc.Status.Code.NOT_FOUND && folderId == null) {
                lines.add("folder /" + folderName + " never created");
            } else {
                clean = false;
                lines.add("folder /" + folderName + " NOT CLEANED UP, " + e.getStatus());
            }
        }
        String lists = "/sites/" + GraphPaths.site(siteId) + "/lists/";
        try {
            String id = listId;
            if (id == null) {
                id = client.get(lists + GraphPaths.segment(listName, "list") + "?$select=id")
                        .path("id").asText("");
            }
            client.delete(lists + GraphPaths.segment(id, "list_id"));
            lines.add("list " + listName + " deleted");
        } catch (GraphClient.GraphApiException e) {
            if (e.status() == 404 && listId == null) {
                lines.add("list " + listName + " never created");
            } else {
                clean = false;
                lines.add("list " + listName + " NOT CLEANED UP, " + e.getMessage());
            }
        } catch (IOException | RuntimeException e) {
            clean = false;
            lines.add("list " + listName + " NOT CLEANED UP, " + e.getMessage());
        }
        check(clean, String.join("; ", lines));
        return String.join("; ", lines);
    }
}
