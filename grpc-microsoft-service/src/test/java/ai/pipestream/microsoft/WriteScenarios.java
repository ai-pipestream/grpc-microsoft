package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ColumnOutcome;
import ai.pipestream.microsoft.v1.ColumnScope;
import ai.pipestream.microsoft.v1.ClearValue;
import ai.pipestream.microsoft.v1.ColumnSpec;
import ai.pipestream.microsoft.v1.HyperlinkValue;
import ai.pipestream.microsoft.v1.MultiChoiceValue;
import ai.pipestream.microsoft.v1.MultiLookupValue;
import ai.pipestream.microsoft.v1.MultiPersonValue;
import ai.pipestream.microsoft.v1.PersonValue;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CopyDriveItemRequest;
import ai.pipestream.microsoft.v1.CopyDriveItemResponse;
import ai.pipestream.microsoft.v1.CreateFolderRequest;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.DateTimeColumnFormat;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeltaGroupsRequest;
import ai.pipestream.microsoft.v1.DeltaGroupsResponse;
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
import ai.pipestream.microsoft.v1.ListPermissionsRequest;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListTemplate;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MutateRequest;
import ai.pipestream.microsoft.v1.MutateResponse;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.RestoreDriveItemVersionRequest;
import ai.pipestream.microsoft.v1.SharePointIdentitySet;
import ai.pipestream.microsoft.v1.UpdateDriveItemRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsRequest;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import ai.pipestream.microsoft.v1.DirectoryMember;
import ai.pipestream.microsoft.v1.GetSiteRoleInheritanceRequest;
import ai.pipestream.microsoft.v1.GraphGroup;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.GroupChange;
import ai.pipestream.microsoft.v1.GroupMemberChange;
import ai.pipestream.microsoft.v1.ListGroupTransitiveMembersRequest;
import ai.pipestream.microsoft.v1.ListGroupsRequest;
import ai.pipestream.microsoft.v1.ListSiteGroupMembersRequest;
import ai.pipestream.microsoft.v1.ListSiteGroupsRequest;
import ai.pipestream.microsoft.v1.ListSiteRoleAssignmentsRequest;
import ai.pipestream.microsoft.v1.ListUserTransitiveMemberOfRequest;
import ai.pipestream.microsoft.v1.ListUsersRequest;
import ai.pipestream.microsoft.v1.SiteGroup;
import ai.pipestream.microsoft.v1.SitePrincipal;
import ai.pipestream.microsoft.v1.SitePrincipalType;
import ai.pipestream.microsoft.v1.SiteRoleAssignment;
import ai.pipestream.microsoft.v1.SiteSecurable;
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
import java.util.HashSet;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
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
    private String userId;
    private String groupId;
    private Integer siteGroupId;
    private String userMail;
    private boolean linkColumn;
    private String libraryId;
    private boolean libraryTried;

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
                new Scenario("list-permissions-drive-item", "ListPermissions",
                        "ListPermissions [uploaded file]", this::permissionsOfTheFile),
                new Scenario("list-permissions-folder", "ListPermissions",
                        "ListPermissions [test folder]", this::permissionsOfTheFolder),
                new Scenario("update-drive-item", "UpdateDriveItem",
                        "UpdateDriveItem [rename]", this::updateDriveItem),
                new Scenario("update-drive-item-description", "UpdateDriveItem",
                        "UpdateDriveItem [description on a SharePoint drive]",
                        this::describeOnASharePointDrive),
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
                new Scenario("list-permissions-list", "ListPermissions",
                        "ListPermissions [test list]", this::permissionsOfTheList),
                new Scenario("list-permissions-list-item", "ListPermissions",
                        "ListPermissions [list item]", this::permissionsOfTheListItem),
                new Scenario("list-users", "ListUsers", "ListUsers", this::users),
                new Scenario("list-groups", "ListGroups", "ListGroups", this::groups),
                new Scenario("delta-groups", "DeltaGroups", "DeltaGroups [full, then from its link]",
                        this::groupChanges),
                new Scenario("list-group-transitive-members", "ListGroupTransitiveMembers",
                        "ListGroupTransitiveMembers", this::membersOfAGroup),
                new Scenario("list-user-transitive-member-of", "ListUserTransitiveMemberOf",
                        "ListUserTransitiveMemberOf", this::groupsOfAUser),
                new Scenario("list-site-groups", "ListSiteGroups", "ListSiteGroups",
                        this::siteGroups),
                new Scenario("list-site-group-members", "ListSiteGroupMembers",
                        "ListSiteGroupMembers", this::siteGroupMembers),
                new Scenario("get-site-role-inheritance-site", "GetSiteRoleInheritance",
                        "GetSiteRoleInheritance [site]", this::inheritanceOfTheSite),
                new Scenario("get-site-role-inheritance-list", "GetSiteRoleInheritance",
                        "GetSiteRoleInheritance [test list]", this::inheritanceOfTheList),
                new Scenario("list-site-role-assignments-site", "ListSiteRoleAssignments",
                        "ListSiteRoleAssignments [site]", this::roleAssignmentsOfTheSite),
                new Scenario("list-site-role-assignments-list", "ListSiteRoleAssignments",
                        "ListSiteRoleAssignments [test list]", this::roleAssignmentsOfTheList),
                new Scenario("get-site-role-inheritance-list-item", "GetSiteRoleInheritance",
                        "GetSiteRoleInheritance [list item]", this::inheritanceOfTheListItem),
                new Scenario("list-site-role-assignments-list-item", "ListSiteRoleAssignments",
                        "ListSiteRoleAssignments [list item]", this::roleAssignmentsOfTheListItem),
                new Scenario("update-list-item-fields", "UpdateListItemFields",
                        "UpdateListItemFields [every column type]", this::updateListItemFields),
                new Scenario("update-list-item-rich-fields", "UpdateListItemFields",
                        "UpdateListItemFields [note, multi choice, currency, link, lookups, clear]",
                        this::updateRichFields),
                new Scenario("update-list-item-person-fields", "UpdateListItemFields",
                        "UpdateListItemFields [person by id, by email, several people]",
                        this::updatePersonFields),
                new Scenario("update-list-item-fields-drive-item", "UpdateListItemFields",
                        "UpdateListItemFields [drive item]", this::updateDriveItemFields),
                new Scenario("mutate", "Mutate", "Mutate", this::mutate),
                new Scenario("links-create-and-mutate", "CreateListItem",
                        "CreateListItem and Mutate [link values]", this::linksThroughCreateAndMutate),
                new Scenario("upload-with-fields", "UploadItem",
                        "UploadItem [text and link fields, into a new library]",
                        this::uploadWithFields),
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
        DriveItem updated = stub().updateDriveItem(UpdateDriveItemRequest.newBuilder()
                .setItem(ref(small.getId())).setExpectedEtag(small.getEtag())
                .setName("renamed.txt").build()).getItem();
        DriveItem read = getItem(small.getId());
        check(read.getName().equals("renamed.txt"), "name read back as " + read.getName());
        check(!updated.getEtag().equals(small.getEtag()), "the eTag did not change");
        return "name read back; eTag changed";
    }

    private String describeOnASharePointDrive() {
        need(small, "the small upload");
        String before = getItem(small.getId()).getEtag();
        try {
            stub().updateDriveItem(UpdateDriveItemRequest.newBuilder().setItem(ref(small.getId()))
                    .setDescription("must not land").build());
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == io.grpc.Status.Code.FAILED_PRECONDITION) {
                check(getItem(small.getId()).getEtag().equals(before), "the item changed");
                return "FAILED_PRECONDITION, item unchanged";
            }
            throw e;
        }
        throw new AssertionError("a description was accepted on a SharePoint drive");
    }

    private String updateWithStaleEtag() {
        need(small, "the small upload");
        check(!getItem(small.getId()).getEtag().equals(small.getEtag()),
                "the item still has its first eTag, so there is no stale one to send");
        try {
            stub().updateDriveItem(UpdateDriveItemRequest.newBuilder().setItem(ref(small.getId()))
                    .setExpectedEtag(small.getEtag()).setName("must-not-land.txt").build());
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == io.grpc.Status.Code.ABORTED) {
                check(!getItem(small.getId()).getName().equals("must-not-land.txt"),
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

    private String updateRichFields() {
        need(itemId, "the list item");
        ColumnSpec.Builder column = ColumnSpec.newBuilder();
        List<String> problems = new ArrayList<>();
        Set<String> created = ensureMoreColumns(problems, List.of(
                column.clone().setName("PsNote").setType(ColumnType.COLUMN_TYPE_NOTE).build(),
                column.clone().setName("PsMulti").setType(ColumnType.COLUMN_TYPE_MULTICHOICE)
                        .addChoices("A").addChoices("B").addChoices("C").build(),
                column.clone().setName("PsMoney").setType(ColumnType.COLUMN_TYPE_CURRENCY).build(),
                column.clone().setName("PsLink").setType(ColumnType.COLUMN_TYPE_URL).build(),
                column.clone().setName("PsLookup").setType(ColumnType.COLUMN_TYPE_LOOKUP)
                        .setLookupListId(listId).build(),
                column.clone().setName("PsLookups").setType(ColumnType.COLUMN_TYPE_LOOKUP)
                        .setLookupListId(listId).setMultiValued(true).build()));
        int self = Integer.parseInt(itemId);
        String id = Integer.toString(self);
        if (created.contains("PsNote")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsNote")
                    .setNote("line one\nline two"), "PsNote", List.of("line one\nline two"));
            writeAndRead(problems, FieldValue.newBuilder().setName("PsNote")
                    .setClear(ClearValue.getDefaultInstance()), "PsNote", List.of());
        }
        if (created.contains("PsMulti")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsMulti").setMultiChoice(
                    MultiChoiceValue.newBuilder().addChoices("A").addChoices("C")),
                    "PsMulti", List.of("A", "C"));
        }
        if (created.contains("PsMoney")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsMoney").setNumber(12.5),
                    "PsMoney", List.of("12.5"));
        }
        linkColumn = created.contains("PsLink");
        if (linkColumn) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsLink").setHyperlink(
                    HyperlinkValue.newBuilder().setUrl("https://example.com/page")
                            .setDescription("A page")),
                    "PsLink.Url", List.of("https://example.com/page"));
        }
        if (created.contains("PsLookup")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsLookup").setLookup(self),
                    "PsLookupLookupId", List.of(id));
        }
        if (created.contains("PsLookups")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsLookups").setMultiLookup(
                    MultiLookupValue.newBuilder().addLookupIds(self)),
                    "PsLookups.LookupId", List.of(id));
        }
        check(problems.isEmpty(), String.join("; ", problems));
        return "note, multi choice, currency, link, lookup and multi lookup read back; clear emptied"
                + " the note";
    }

    private String updatePersonFields() {
        need(itemId, "the list item");
        need(userMail, "a directory user with a mail address");
        ColumnSpec.Builder column = ColumnSpec.newBuilder().setType(ColumnType.COLUMN_TYPE_USER);
        List<String> problems = new ArrayList<>();
        Set<String> created = ensureMoreColumns(problems, List.of(
                column.clone().setName("PsOwner").build(),
                column.clone().setName("PsPerson").build(),
                column.clone().setName("PsPeople").setMultiValued(true).build()));
        if (created.contains("PsOwner")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsOwner").setPerson(
                    PersonValue.newBuilder().setEmail(userMail)), "PsOwnerLookupId", null);
        }
        String id = fields(itemId).getOrDefault("PsOwnerLookupId", "");
        if (id.isEmpty()) {
            problems.add("no person id to write the other person columns with");
        } else {
            int person = Integer.parseInt(id);
            if (created.contains("PsPerson")) {
                writeAndRead(problems, FieldValue.newBuilder().setName("PsPerson").setPerson(
                        PersonValue.newBuilder().setLookupId(person)),
                        "PsPersonLookupId", List.of(id));
            }
            if (created.contains("PsPeople")) {
                writeAndRead(problems, FieldValue.newBuilder().setName("PsPeople").setMultiPerson(
                        MultiPersonValue.newBuilder().addPeople(
                                PersonValue.newBuilder().setLookupId(person))),
                        "PsPeople.LookupId", List.of(id));
            }
        }
        check(problems.isEmpty(), String.join("; ", problems));
        return "person by email, person by id and a people column read back";
    }

    /** Creates columns on the test list one by one; notes a problem for each that is not created. */
    private Set<String> ensureMoreColumns(List<String> problems, List<ColumnSpec> columns) {
        need(listId, "the test list");
        Set<String> created = new HashSet<>();
        for (ColumnSpec column : columns) {
            try {
                ColumnOutcome outcome = stub().ensureColumns(EnsureColumnsRequest.newBuilder()
                        .setScope(ColumnScope.newBuilder()
                                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId)))
                        .addColumns(column).build()).getColumns(0);
                if (outcome.getState() == EnsureState.ENSURE_STATE_CREATED) {
                    created.add(column.getName());
                } else {
                    problems.add(column.getName() + " column answered " + outcome.getState() + " "
                            + outcome.getError().getMessage());
                }
            } catch (StatusRuntimeException e) {
                problems.add(column.getName() + " column refused: " + e.getStatus().getCode() + " "
                        + e.getStatus().getDescription());
            }
        }
        return created;
    }

    /** Writes one column and notes a problem when the write fails or reads back differently. */
    private void writeAndRead(List<String> problems, FieldValue.Builder value, String readName,
            List<String> expected) {
        String label = value.getName() + " (" + value.getValueCase() + ")";
        try {
            stub().updateListItemFields(UpdateListItemFieldsRequest.newBuilder()
                    .setExpectedEtag(listItem(itemId).getEtag())
                    .setPatch(ListItemPatch.newBuilder().setTarget(listItemTarget(itemId))
                            .addFields(value))
                    .build());
        } catch (StatusRuntimeException e) {
            problems.add(label + " refused: " + e.getStatus().getCode() + " "
                    + e.getStatus().getDescription());
            return;
        }
        List<String> read = List.of();
        for (ListColumn column : listItem(itemId).getFieldsList()) {
            if (column.getName().equals(readName)) {
                read = column.getStringValuesCount() > 0 ? column.getStringValuesList()
                        : List.of(fields(itemId).get(readName));
            }
        }
        if (expected == null ? read.isEmpty() : !read.equals(expected)) {
            problems.add(label + " read back as " + readName + "=" + read);
        }
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
                MutateRequest.newBuilder().setOpId("rename").setUpdateDriveItem(
                        UpdateDriveItemRequest.newBuilder().setItem(ref(copy.getId()))
                                .setName("renamed-in-batch.txt")).build(),
                MutateRequest.newBuilder().setOpId("drop-copy").addDependsOn("rename")
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

    private String linksThroughCreateAndMutate() throws Exception {
        need(itemId, "the list item");
        if (!linkColumn) {
            throw new Skip("needs the link column, which an earlier step did not produce");
        }
        ListRef list = ListRef.newBuilder().setSiteId(siteId).setListId(listId).build();
        ListItem created = stub().createListItem(CreateListItemRequest.newBuilder().setList(list)
                .addFields(FieldValue.newBuilder().setName("Title").setText("with link"))
                .addFields(link("https://example.com/created")).build()).getItem();
        check(created.getFieldsList().stream().anyMatch(column -> column.getName()
                .equals("PsLink.Url") && column.getStringValue().equals("https://example.com/created")),
                "CreateListItem answered without the link");
        check("https://example.com/created".equals(fields(created.getId()).get("PsLink.Url")),
                "the created item's link reads back differently");
        List<MutateRequest> ops = List.of(
                MutateRequest.newBuilder().setOpId("new-with-link").setCreateListItem(
                        CreateListItemRequest.newBuilder().setList(list)
                                .addFields(FieldValue.newBuilder().setName("Title").setText("batch link"))
                                .addFields(link("https://example.com/batch-created"))).build(),
                MutateRequest.newBuilder().setOpId("relink").setUpdateListItemFields(
                        UpdateListItemFieldsRequest.newBuilder().setPatch(ListItemPatch.newBuilder()
                                .setTarget(listItemTarget(itemId))
                                .addFields(link("https://example.com/batch-updated")))).build());
        Collector<MutateResponse> responses = new Collector<>();
        StreamObserver<MutateRequest> requests = async.mutate(responses);
        ops.forEach(requests::onNext);
        requests.onCompleted();
        Map<String, MutateResponse> byId = new HashMap<>();
        responses.await().forEach(response -> byId.put(response.getOpId(), response));
        check(byId.size() == ops.size(), "answered " + byId.size() + " of " + ops.size() + " ops");
        for (MutateResponse response : byId.values()) {
            check(!response.hasError(), "op " + response.getOpId() + " "
                    + response.getError().getCode() + " " + response.getError().getMessage());
        }
        check("https://example.com/batch-created".equals(fields(byId.get("new-with-link")
                .getCreateListItem().getItem().getId()).get("PsLink.Url")),
                "the batch-created item's link reads back differently");
        check("https://example.com/batch-updated".equals(fields(itemId).get("PsLink.Url")),
                "the batch-updated link reads back differently");
        return "a link written by CreateListItem and by two Mutate ops, each read back";
    }

    private String uploadWithFields() throws Exception {
        String libraryName = listName + "-lib";
        libraryTried = true;
        EnsureListResponse library = stub().ensureList(EnsureListRequest.newBuilder()
                .setSiteId(siteId).setDisplayName(libraryName)
                .setDescription("Created by the live write test; safe to delete.")
                .setTemplate(ListTemplate.LIST_TEMPLATE_DOCUMENT_LIBRARY).build());
        check(library.getState() == EnsureState.ENSURE_STATE_CREATED,
                "the library answered " + library.getState());
        libraryId = library.getList().getId();
        ColumnSpec.Builder column = ColumnSpec.newBuilder();
        for (ColumnOutcome outcome : stub().ensureColumns(EnsureColumnsRequest.newBuilder()
                .setScope(ColumnScope.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(siteId).setListId(libraryId)))
                .addColumns(column.clone().setName("PsText").setType(ColumnType.COLUMN_TYPE_TEXT))
                .addColumns(column.clone().setName("PsLink").setType(ColumnType.COLUMN_TYPE_URL))
                .build()).getColumnsList()) {
            check(outcome.getState() == EnsureState.ENSURE_STATE_CREATED, outcome.getName()
                    + " answered " + outcome.getState() + " " + outcome.getError().getMessage());
        }
        String libraryDrive = client.get(GraphPaths.list(siteId, libraryId) + "/drive?$select=id")
                .path("id").asText("");
        check(!libraryDrive.isEmpty(), "the library has no drive");
        byte[] content = remember(("with fields " + folderName + "\n")
                .getBytes(StandardCharsets.UTF_8));
        Collector<UploadItemResponse> responses = new Collector<>();
        StreamObserver<UploadItemRequest> requests = async.uploadItem(responses);
        requests.onNext(frame(UploadItemHeader.newBuilder().setDriveId(libraryDrive)
                .setName("with-fields.txt").setSize(content.length)
                .addFields(FieldValue.newBuilder().setName("PsText").setText("from the upload"))
                .addFields(link("https://example.com/uploaded"))));
        sendChunks(requests, content, 0, content.length);
        requests.onNext(end(sha256(content)));
        requests.onCompleted();
        List<UploadItemResponse> all = responses.await();
        Map<String, String> columns = new HashMap<>();
        all.get(all.size() - 1).getItem().getListColumnsList()
                .forEach(c -> columns.put(c.getName(), c.getStringValue()));
        check("from the upload".equals(columns.get("PsText")),
                "PsText answered as " + columns.get("PsText"));
        check("https://example.com/uploaded".equals(columns.get("PsLink.Url")),
                "PsLink answered as " + columns.get("PsLink.Url"));
        return "a file uploaded into a new library with a text and a link column, both answered";
    }

    private static FieldValue.Builder link(String url) {
        return FieldValue.newBuilder().setName("PsLink")
                .setHyperlink(HyperlinkValue.newBuilder().setUrl(url).setDescription("A page"));
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

    // ------------------------------------------------------------ directory

    private String users() {
        List<GraphUser> users = new ArrayList<>();
        stub().listUsers(ListUsersRequest.getDefaultInstance())
                .forEachRemaining(response -> users.add(response.getUser()));
        check(!users.isEmpty(), "the directory has no user");
        check(users.stream().noneMatch(user -> user.getId().isEmpty()), "a user has no id");
        check(users.stream().map(GraphUser::getId).distinct().count() == users.size(),
                "a user was listed twice");
        check(users.stream().anyMatch(user -> !user.getUserPrincipalName().isEmpty()),
                "no user has a sign-in name");
        userId = users.get(0).getId();
        userMail = users.stream().map(GraphUser::getMail).filter(mail -> !mail.isEmpty())
                .findFirst().orElse(null);
        return users.size() + " user(s)";
    }

    private String groups() {
        List<GraphGroup> groups = new ArrayList<>();
        stub().listGroups(ListGroupsRequest.getDefaultInstance())
                .forEachRemaining(response -> groups.add(response.getGroup()));
        check(groups.stream().noneMatch(group -> group.getId().isEmpty()), "a group has no id");
        check(groups.stream().map(GraphGroup::getId).distinct().count() == groups.size(),
                "a group was listed twice");
        if (!groups.isEmpty()) {
            groupId = groups.get(0).getId();
        }
        return groups.size() + " group(s)";
    }

    private String groupChanges() {
        List<DeltaGroupsResponse> full = new ArrayList<>();
        stub().deltaGroups(DeltaGroupsRequest.newBuilder().setIncludeMembers(true).build())
                .forEachRemaining(full::add);
        check(!full.isEmpty() && full.get(full.size() - 1).getEventCase()
                == DeltaGroupsResponse.EventCase.DELTA_LINK, "the stream did not end with a delta link");
        String link = full.get(full.size() - 1).getDeltaLink();
        check(!link.isEmpty(), "the delta link is empty");
        Set<String> ids = new TreeSet<>();
        int members = 0;
        for (DeltaGroupsResponse response : full.subList(0, full.size() - 1)) {
            check(response.hasChange(), "a full enumeration held something other than a change");
            GroupChange change = response.getChange();
            check(!change.getGroup().getId().isEmpty(), "a group change has no group id");
            if (!change.getRemoved()) {
                ids.add(change.getGroup().getId());
            }
            for (GroupMemberChange member : change.getMemberChangesList()) {
                check(!member.getId().isEmpty() && !member.getType().isEmpty(),
                        "a member change has no id or type");
                members++;
            }
        }
        check(groupId == null || ids.contains(groupId), "a listed group is missing from the delta");

        List<DeltaGroupsResponse> next = new ArrayList<>();
        stub().deltaGroups(DeltaGroupsRequest.newBuilder().setDeltaLink(link).setIncludeMembers(true)
                .build()).forEachRemaining(next::add);
        check(!next.isEmpty() && !next.get(next.size() - 1).getDeltaLink().isEmpty(),
                "the second stream did not end with a delta link");
        check(next.stream().noneMatch(DeltaGroupsResponse::hasResync), "the fresh link was refused");
        return ids.size() + " group(s) with " + members + " member change(s); from the link, "
                + (next.size() - 1) + " change(s)";
    }

    private String membersOfAGroup() {
        need(groupId, "a group");
        List<DirectoryMember> members = new ArrayList<>();
        stub().listGroupTransitiveMembers(ListGroupTransitiveMembersRequest.newBuilder()
                .setGroupId(groupId).build())
                .forEachRemaining(response -> members.add(response.getMember()));
        return members.size() + " member(s), kinds " + memberKinds(members);
    }

    private String groupsOfAUser() {
        need(userId, "a user");
        List<DirectoryMember> memberships = new ArrayList<>();
        stub().listUserTransitiveMemberOf(ListUserTransitiveMemberOfRequest.newBuilder()
                .setUserId(userId).build())
                .forEachRemaining(response -> memberships.add(response.getMember()));
        check(memberships.stream().noneMatch(DirectoryMember::hasUser), "a user is a member of a user");
        return memberships.size() + " membership(s), kinds " + memberKinds(memberships);
    }

    /** How many members there are of each kind; fails when one has no id. */
    private static Map<String, Integer> memberKinds(List<DirectoryMember> members) {
        Map<String, Integer> kinds = new TreeMap<>();
        for (DirectoryMember member : members) {
            String id = switch (member.getMemberCase()) {
                case USER -> member.getUser().getId();
                case GROUP -> member.getGroup().getId();
                case OTHER -> member.getOther().getId();
                case MEMBER_NOT_SET -> "";
            };
            check(!id.isEmpty(), "a member has no id");
            kinds.merge(member.getMemberCase() == DirectoryMember.MemberCase.OTHER
                    ? member.getOther().getType() : member.getMemberCase().name().toLowerCase(
                            java.util.Locale.ROOT), 1, Integer::sum);
        }
        return kinds;
    }

    // ------------------------------------------------------------ site security

    private String siteGroups() {
        List<SiteGroup> groups = new ArrayList<>();
        stub().listSiteGroups(ListSiteGroupsRequest.newBuilder().setSiteId(siteId).build())
                .forEachRemaining(response -> groups.add(response.getGroup()));
        check(!groups.isEmpty(), "the site has no group");
        for (SiteGroup group : groups) {
            check(group.getId() > 0, "a site group has no id");
            check(!group.getTitle().isEmpty(), "a site group has no title");
        }
        check(groups.stream().map(SiteGroup::getId).distinct().count() == groups.size(),
                "a site group was listed twice");
        siteGroupId = groups.get(0).getId();
        return groups.size() + " group(s)";
    }

    private String siteGroupMembers() {
        need(siteGroupId, "a site group");
        List<SitePrincipal> members = new ArrayList<>();
        try {
            stub().listSiteGroupMembers(ListSiteGroupMembersRequest.newBuilder().setSiteId(siteId)
                    .setGroupId(siteGroupId).build())
                    .forEachRemaining(response -> members.add(response.getMember()));
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        Map<String, Integer> kinds = new TreeMap<>();
        for (SitePrincipal member : members) {
            checkPrincipal(member);
            kinds.merge(member.getPrincipalType().name(), 1, Integer::sum);
        }
        return members.size() + " member(s), kinds " + kinds;
    }

    private String inheritanceOfTheSite() {
        boolean unique = stub().getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setSiteId(siteId)).build())
                .getHasUniqueRoleAssignments();
        check(unique, "the site's root web has no role assignments of its own");
        return "the site has its own role assignments";
    }

    private String inheritanceOfTheList() {
        need(listId, "the test list");
        boolean unique = stub().getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(listSecurable()).build()).getHasUniqueRoleAssignments();
        check(!unique, "a new list does not inherit the site's role assignments");
        return "the new list inherits from the site";
    }

    private String roleAssignmentsOfTheSite() {
        return roleAssignments(SiteSecurable.newBuilder().setSiteId(siteId).build());
    }

    private String roleAssignmentsOfTheList() {
        need(listId, "the test list");
        return roleAssignments(listSecurable());
    }

    private String inheritanceOfTheListItem() {
        need(itemId, "the list item");
        boolean unique = stub().getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(listItemRef(itemId))).build())
                .getHasUniqueRoleAssignments();
        check(!unique, "a new list item does not inherit the list's role assignments");
        return "the new list item inherits from the list";
    }

    private String roleAssignmentsOfTheListItem() {
        need(itemId, "the list item");
        return roleAssignments(SiteSecurable.newBuilder().setListItem(listItemRef(itemId)).build());
    }

    private SiteSecurable listSecurable() {
        return SiteSecurable.newBuilder()
                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId)).build();
    }

    /** Lists the role assignments of a securable and answers how many and their permission levels. */
    private String roleAssignments(SiteSecurable securable) {
        List<SiteRoleAssignment> assignments = new ArrayList<>();
        try {
            stub().listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                    .setSecurable(securable).build())
                    .forEachRemaining(response -> assignments.add(response.getAssignment()));
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        check(!assignments.isEmpty(), "nothing is assigned");
        Set<String> levels = new TreeSet<>();
        for (SiteRoleAssignment assignment : assignments) {
            check(assignment.hasMember(), "an assignment names no member");
            check(assignment.getPrincipalId() == assignment.getMember().getId(),
                    "an assignment's member is not its principal");
            checkPrincipal(assignment.getMember());
            check(assignment.getRolesCount() > 0, "an assignment has no permission level");
            assignment.getRolesList().forEach(role -> {
                check(role.getId() > 0 && !role.getName().isEmpty(), "a permission level has no id or name");
                levels.add(role.getName());
            });
        }
        return assignments.size() + " assignment(s), levels " + levels;
    }

    private static void checkPrincipal(SitePrincipal principal) {
        check(principal.getId() > 0, "a principal has no id");
        check(principal.getPrincipalType() != SitePrincipalType.SITE_PRINCIPAL_TYPE_UNSPECIFIED,
                "a principal has no type");
        check(!principal.getLoginName().isEmpty(), "a principal has no login name");
    }

    /** A skip when SharePoint refused the read for lack of permission; any other error as it is. */
    private static RuntimeException notGranted(StatusRuntimeException e) {
        if (e.getStatus().getCode() != io.grpc.Status.Code.PERMISSION_DENIED) {
            return e;
        }
        return new Skip("the app's grant on the site does not allow this read: "
                + e.getStatus().getDescription());
    }

    private String permissionsOfTheFile() {
        need(small, "the small upload");
        return permissions(PermissionTarget.newBuilder().setDriveItem(ref(small.getId())).build());
    }

    private String permissionsOfTheFolder() {
        need(folderId, "the test folder");
        return permissions(PermissionTarget.newBuilder().setDriveItem(ref(folderId)).build());
    }

    private String permissionsOfTheList() {
        need(listId, "the test list");
        return permissions(PermissionTarget.newBuilder()
                .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId)).build());
    }

    private String permissionsOfTheListItem() {
        need(itemId, "the list item");
        return permissions(PermissionTarget.newBuilder().setListItem(listItemRef(itemId)).build());
    }

    /** Lists the ACL of a target and answers how many entries, their grantee kinds and roles. */
    private String permissions(PermissionTarget target) {
        List<Permission> permissions = new ArrayList<>();
        stub().listPermissions(ListPermissionsRequest.newBuilder().setTarget(target).build())
                .forEachRemaining(response -> permissions.add(response.getPermission()));
        check(!permissions.isEmpty(), "the ACL is empty");
        Map<String, Integer> kinds = new TreeMap<>();
        Set<String> roles = new TreeSet<>();
        for (Permission permission : permissions) {
            check(!permission.getId().isEmpty(), "a permission has no id");
            check(permission.getRolesCount() > 0, "a permission has no role");
            check(permission.getTarget().equals(target), "a permission names another target");
            roles.addAll(permission.getRolesList());
            List<SharePointIdentitySet> grantees = new ArrayList<>(
                    permission.getGrantedToIdentitiesList());
            if (permission.hasGrantedTo()) {
                grantees.add(permission.getGrantedTo());
            }
            check(!grantees.isEmpty() || permission.hasLink(), "a permission names no grantee");
            if (permission.hasLink()) {
                kinds.merge("link", 1, Integer::sum);
            }
            grantees.forEach(grantee -> kinds.merge(kind(grantee), 1, Integer::sum));
        }
        return permissions.size() + " permission(s), grantees " + kinds + ", roles " + roles;
    }

    private static String kind(SharePointIdentitySet grantee) {
        if (grantee.hasSharePointGroup() || grantee.hasSiteGroup()) {
            return "site group";
        }
        if (grantee.hasGroup()) {
            return "group";
        }
        if (grantee.hasUser() || grantee.hasSiteUser()) {
            return "user";
        }
        if (grantee.hasApplication()) {
            return "application";
        }
        return grantee.hasDevice() ? "device" : "unknown";
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
        String libraryName = listName + "-lib";
        try {
            if (!libraryTried) {
                return finish(clean, lines);
            }
            String id = libraryId;
            if (id == null) {
                id = client.get(lists + GraphPaths.segment(libraryName, "list") + "?$select=id")
                        .path("id").asText("");
            }
            client.delete(lists + GraphPaths.segment(id, "list_id"));
            lines.add("library " + libraryName + " deleted");
        } catch (GraphClient.GraphApiException e) {
            if (e.status() != 404 || libraryId != null) {
                clean = false;
                lines.add("library " + libraryName + " NOT CLEANED UP, " + e.getMessage());
            }
        } catch (IOException | RuntimeException e) {
            clean = false;
            lines.add("library " + libraryName + " NOT CLEANED UP, " + e.getMessage());
        }
        return finish(clean, lines);
    }

    private static String finish(boolean clean, List<String> lines) {
        check(clean, String.join("; ", lines));
        return String.join("; ", lines);
    }
}
