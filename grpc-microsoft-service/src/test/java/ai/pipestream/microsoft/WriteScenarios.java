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
import ai.pipestream.microsoft.v1.CreateSharingLinkRequest;
import ai.pipestream.microsoft.v1.CreateSharingLinkResponse;
import ai.pipestream.microsoft.v1.DeletePermissionRequest;
import ai.pipestream.microsoft.v1.GetPermissionRequest;
import ai.pipestream.microsoft.v1.GrantLinkAccessRequest;
import ai.pipestream.microsoft.v1.InviteRecipientsRequest;
import ai.pipestream.microsoft.v1.RevokeLinkAccessRequest;
import ai.pipestream.microsoft.v1.SharingRecipient;
import ai.pipestream.microsoft.v1.SiteSharingLinkGroup;
import ai.pipestream.microsoft.v1.UpdatePermissionRequest;
import ai.pipestream.microsoft.v1.DateTimeColumnFormat;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeltaGroupsRequest;
import ai.pipestream.microsoft.v1.DeltaGroupsResponse;
import ai.pipestream.microsoft.v1.DeltaUsersRequest;
import ai.pipestream.microsoft.v1.DeltaUsersResponse;
import ai.pipestream.microsoft.v1.ListSiteChangesRequest;
import ai.pipestream.microsoft.v1.ListSiteChangesResponse;
import ai.pipestream.microsoft.v1.SiteChange;
import ai.pipestream.microsoft.v1.SiteChangeObject;
import ai.pipestream.microsoft.v1.SiteChangeScope;
import ai.pipestream.microsoft.v1.SiteChangeType;
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
import java.util.EnumMap;
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
     * @param termSetId the term set that term columns draw from; empty when there is none
     * @param termIds ids of the terms {@link #TERM_LABELS} of that set, in that order
     */
    record Backend(String siteId, String driveId, String suffix, Duration copyPollInterval,
            Map<String, byte[]> blobs, String termSetId, List<String> termIds) {
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
    private SiteSecurable folderSecurable;
    private boolean inheritanceBroken;
    /** The resume cursor of the security sync that ran while inheritance was broken. */
    private String securityCursor;
    /** The end of each change log before the list item and folder stopped inheriting. */
    private final Map<SiteChangeScope, String> changeTokens = new EnumMap<>(SiteChangeScope.class);
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
    private String userPrincipalName;
    /** The SharePoint group this run created and has not deleted yet. */
    private Integer ownGroupId;
    private Integer ownGroupMemberId;
    /** The group this run created, kept after its deletion. */
    private Integer createdGroupId;
    /** The Entra group this run created and has not deleted yet. */
    private String entraGroupId;
    private Long readRoleId;
    /** The file this run shares. */
    private DriveItem shared;
    private SiteSecurable sharedSecurable;
    private String organizationLinkId;
    private String usersLinkId;
    private String anyoneLinkId;
    private String invitedPermissionId;
    /** Roles the invited group held on the shared file before the invitation. */
    private final Set<String> rolesBeforeTheInvitation = new TreeSet<>();
    private boolean linkColumn;
    private String libraryId;
    private boolean libraryTried;
    private GraphGroup listedGroup;
    private final Map<String, List<com.google.protobuf.Message>> records = new TreeMap<>();

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

    /**
     * What the directory and security reads of this run answered, by kind:
     * {@code entra-users}, {@code entra-group-changes}, {@code entra-group-owners}
     * (the group, then its owners) and {@code sharepoint-sync-records}.
     */
    Map<String, List<com.google.protobuf.Message>> records() {
        return records;
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
                new Scenario("delta-users", "DeltaUsers", "DeltaUsers [full, then from its link]",
                        this::userChanges),
                new Scenario("list-group-transitive-members", "ListGroupTransitiveMembers",
                        "ListGroupTransitiveMembers", this::membersOfAGroup),
                new Scenario("list-group-owners", "ListGroupOwners", "ListGroupOwners",
                        this::ownersOfAGroup),
                new Scenario("list-user-transitive-member-of", "ListUserTransitiveMemberOf",
                        "ListUserTransitiveMemberOf", this::groupsOfAUser),
                new Scenario("create-group", "CreateGroup", "CreateGroup [a security group]",
                        this::createGroup),
                new Scenario("get-group", "GetGroup", "GetGroup [the new group]", this::getGroup),
                new Scenario("update-group", "UpdateGroup", "UpdateGroup [name and description]",
                        this::updateGroup),
                new Scenario("add-group-member", "AddGroupMember", "AddGroupMember [a directory user]",
                        this::addGroupMember),
                new Scenario("remove-group-member", "RemoveGroupMember", "RemoveGroupMember",
                        this::removeGroupMember),
                new Scenario("delete-group", "DeleteGroup", "DeleteGroup", this::deleteGroup),
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
                new Scenario("list-site-changes-latest", "ListSiteChanges",
                        "ListSiteChanges [the end of the site, site collection and list logs]",
                        this::endOfTheChangeLogs),
                new Scenario("create-sharing-link-organization", "CreateSharingLink",
                        "CreateSharingLink [a view link for the organization, asked for twice]",
                        this::shareWithTheOrganization),
                new Scenario("get-permission", "GetPermission", "GetPermission [the organization link]",
                        this::getPermission),
                new Scenario("create-sharing-link-users", "CreateSharingLink",
                        "CreateSharingLink [a view link for named people]", this::shareWithNamedPeople),
                new Scenario("grant-link-access", "GrantLinkAccess",
                        "GrantLinkAccess [a directory user on the link for named people]",
                        this::grantLinkAccess),
                new Scenario("invite-recipients", "InviteRecipients",
                        "InviteRecipients [a directory group, no mail]", this::inviteRecipients),
                new Scenario("update-permission", "UpdatePermission",
                        "UpdatePermission [read to write]", this::updatePermission),
                new Scenario("create-sharing-link-anonymous", "CreateSharingLink",
                        "CreateSharingLink [a view link for anyone]", this::shareWithAnyone),
                new Scenario("list-permissions-shared-item", "ListPermissions",
                        "ListPermissions [the shared file]", this::permissionsOfTheSharedFile),
                new Scenario("list-site-role-assignments-shared-item", "ListSiteRoleAssignments",
                        "ListSiteRoleAssignments [the shared file, its link groups and their members]",
                        this::roleAssignmentsOfTheSharedFile),
                new Scenario("set-site-role-inheritance-break", "SetSiteRoleInheritance",
                        "SetSiteRoleInheritance [list item and folder stop inheriting]",
                        this::breakInheritance),
                new Scenario("sync-security", "Sync",
                        "Sync [site security: groups, role assignments, owner of every item]",
                        this::syncSecurity),
                new Scenario("revoke-link-access", "RevokeLinkAccess",
                        "RevokeLinkAccess [the directory user]", this::revokeLinkAccess),
                new Scenario("delete-permission", "DeletePermission",
                        "DeletePermission [the grant and every link of the shared file]",
                        this::deletePermissions),
                new Scenario("create-site-group", "CreateSiteGroup", "CreateSiteGroup",
                        this::createSiteGroup),
                new Scenario("list-site-role-definitions", "ListSiteRoleDefinitions",
                        "ListSiteRoleDefinitions", this::roleDefinitions),
                new Scenario("add-site-group-member", "AddSiteGroupMember",
                        "AddSiteGroupMember [a directory user, by user principal name]",
                        this::addSiteGroupMember),
                new Scenario("add-site-role-assignment", "AddSiteRoleAssignment",
                        "AddSiteRoleAssignment [the new group may read the list item]",
                        this::addRoleAssignment),
                new Scenario("remove-site-role-assignment", "RemoveSiteRoleAssignment",
                        "RemoveSiteRoleAssignment", this::removeRoleAssignment),
                new Scenario("remove-site-group-member", "RemoveSiteGroupMember",
                        "RemoveSiteGroupMember", this::removeSiteGroupMember),
                new Scenario("delete-site-group", "DeleteSiteGroup", "DeleteSiteGroup",
                        this::deleteSiteGroup),
                new Scenario("set-site-role-inheritance-reset", "SetSiteRoleInheritance",
                        "SetSiteRoleInheritance [list item and folder inherit again]",
                        this::restoreInheritance),
                new Scenario("sync-security-resumed", "Sync",
                        "Sync [resumed after the reset: what inherits again arrives with its new owner]",
                        this::syncSecurityResumed),
                new Scenario("list-site-changes", "ListSiteChanges",
                        "ListSiteChanges [what the break and the reset left in each log]",
                        this::siteChanges),
                new Scenario("list-site-changes-old-token", "ListSiteChanges",
                        "ListSiteChanges [a token older than the change log]",
                        this::siteChangesFromAnOldToken),
                new Scenario("list-site-changes-bad-token", "ListSiteChanges",
                        "ListSiteChanges [a token that is not one]",
                        this::siteChangesFromABadToken),
                new Scenario("update-list-item-fields", "UpdateListItemFields",
                        "UpdateListItemFields [every column type]", this::updateListItemFields),
                new Scenario("update-list-item-rich-fields", "UpdateListItemFields",
                        "UpdateListItemFields [note, multi choice, currency, link, lookups, clear]",
                        this::updateRichFields),
                new Scenario("update-list-item-person-fields", "UpdateListItemFields",
                        "UpdateListItemFields [person by id, by email, several people]",
                        this::updatePersonFields),
                new Scenario("term-store", "ListTerms",
                        "GetTermStore, ListTermGroups, ListTermSets, ListTerms [test term set]",
                        this::termStore),
                new Scenario("update-list-item-term-fields", "UpdateListItemFields",
                        "UpdateListItemFields [term, several terms]", this::updateTermFields),
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

    /** Name of the term group the live tenant keeps for this test. */
    static final String TERM_GROUP = "pipestream-live-test";

    /** Default labels of the two terms the test writes, in the order of {@link Backend#termIds}. */
    static final List<String> TERM_LABELS = List.of("Engineering", "Research");

    private List<ai.pipestream.microsoft.v1.TermValue> terms() {
        if (backend.termSetId().isBlank() || backend.termIds().size() != TERM_LABELS.size()) {
            throw new Skip("needs a term set and its " + TERM_LABELS.size() + " test terms");
        }
        List<ai.pipestream.microsoft.v1.TermValue> terms = new ArrayList<>();
        for (int i = 0; i < TERM_LABELS.size(); i++) {
            terms.add(ai.pipestream.microsoft.v1.TermValue.newBuilder()
                    .setLabel(TERM_LABELS.get(i)).setTermId(backend.termIds().get(i)).build());
        }
        return terms;
    }

    private String termStore() {
        List<ai.pipestream.microsoft.v1.TermValue> terms = terms();
        String termSet = backend.termSetId();
        ai.pipestream.microsoft.v1.TermStore store;
        try {
            store = stub().getTermStore(ai.pipestream.microsoft.v1.GetTermStoreRequest.newBuilder()
                    .setSiteId(siteId).build()).getStore();
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == io.grpc.Status.Code.NOT_FOUND) {
                throw new Skip("Graph shows an application token no term store: "
                        + e.getStatus().getDescription());
            }
            throw notGranted(e);
        }
        check(!store.getId().isEmpty(), "the term store has no id");
        check(!store.getDefaultLanguageTag().isEmpty(), "the term store has no default language");
        List<ai.pipestream.microsoft.v1.TermGroup> groups = new ArrayList<>();
        stub().listTermGroups(ai.pipestream.microsoft.v1.ListTermGroupsRequest.newBuilder()
                .setSiteId(siteId).build()).forEachRemaining(r -> groups.add(r.getGroup()));
        String groupId = "";
        for (ai.pipestream.microsoft.v1.TermGroup group : groups) {
            check(!group.getId().isEmpty(), "a term group has no id");
            if (group.getDisplayName().equals(TERM_GROUP)) {
                groupId = group.getId();
            }
        }
        check(!groupId.isEmpty(), "the term store has no group named " + TERM_GROUP);
        List<ai.pipestream.microsoft.v1.TermSet> sets = new ArrayList<>();
        stub().listTermSets(ai.pipestream.microsoft.v1.ListTermSetsRequest.newBuilder()
                .setSiteId(siteId).setGroupId(groupId).build())
                .forEachRemaining(r -> sets.add(r.getSet()));
        check(sets.stream().anyMatch(set -> set.getId().equals(termSet)),
                "the test term group does not hold the test term set");
        List<ai.pipestream.microsoft.v1.Term> found = new ArrayList<>();
        stub().listTerms(ai.pipestream.microsoft.v1.ListTermsRequest.newBuilder()
                .setSiteId(siteId).setSetId(termSet).build())
                .forEachRemaining(r -> found.add(r.getTerm()));
        for (ai.pipestream.microsoft.v1.TermValue wanted : terms) {
            check(found.stream().anyMatch(term -> term.getId().equals(wanted.getTermId())
                    && term.getLabelsList().stream().anyMatch(label -> label.getIsDefault()
                            && label.getName().equals(wanted.getLabel()))),
                    "the test term set does not list " + wanted.getLabel());
        }
        return groups.size() + " group(s), " + sets.size() + " set(s) in the test group, "
                + found.size() + " term(s) in the test set";
    }

    private String updateTermFields() {
        need(itemId, "the list item");
        List<ai.pipestream.microsoft.v1.TermValue> terms = terms();
        ColumnSpec.Builder column = ColumnSpec.newBuilder().setTermSetId(backend.termSetId());
        List<String> problems = new ArrayList<>();
        Set<String> created = ensureMoreColumns(problems, List.of(
                column.clone().setName("PsTerm").setType(ColumnType.COLUMN_TYPE_TERM).build(),
                column.clone().setName("PsTerms").setType(ColumnType.COLUMN_TYPE_MULTITERM).build()));
        ai.pipestream.microsoft.v1.TermValue first = terms.get(0);
        ai.pipestream.microsoft.v1.TermValue second = terms.get(1);
        if (created.contains("PsTerm")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsTerm").setTerm(first),
                    "PsTerm.TermGuid", List.of(first.getTermId()));
        }
        if (created.contains("PsTerms")) {
            writeAndRead(problems, FieldValue.newBuilder().setName("PsTerms").setMultiTerm(
                    ai.pipestream.microsoft.v1.MultiTermValue.newBuilder()
                            .addTerms(first).addTerms(second)),
                    "PsTerms.TermGuid", List.of(first.getTermId(), second.getTermId()));
        }
        check(problems.isEmpty(), String.join("; ", problems));
        return "a term and two terms read back";
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
        userPrincipalName = users.stream().map(GraphUser::getUserPrincipalName)
                .filter(name -> !name.isEmpty()).findFirst().orElse(null);
        records.put("entra-users", List.copyOf(users));
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
            listedGroup = groups.get(0);
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
        records.put("entra-group-changes", full.subList(0, full.size() - 1).stream()
                .<com.google.protobuf.Message>map(DeltaGroupsResponse::getChange).toList());

        List<DeltaGroupsResponse> next = new ArrayList<>();
        stub().deltaGroups(DeltaGroupsRequest.newBuilder().setDeltaLink(link).setIncludeMembers(true)
                .build()).forEachRemaining(next::add);
        check(!next.isEmpty() && !next.get(next.size() - 1).getDeltaLink().isEmpty(),
                "the second stream did not end with a delta link");
        check(next.stream().noneMatch(DeltaGroupsResponse::hasResync), "the fresh link was refused");
        return ids.size() + " group(s) with " + members + " member change(s); from the link, "
                + (next.size() - 1) + " change(s)";
    }

    private String userChanges() {
        List<DeltaUsersResponse> full = new ArrayList<>();
        stub().deltaUsers(DeltaUsersRequest.getDefaultInstance()).forEachRemaining(full::add);
        check(!full.isEmpty() && full.get(full.size() - 1).getEventCase()
                == DeltaUsersResponse.EventCase.DELTA_LINK, "the stream did not end with a delta link");
        String link = full.get(full.size() - 1).getDeltaLink();
        check(!link.isEmpty(), "the delta link is empty");
        Set<String> ids = new TreeSet<>();
        for (DeltaUsersResponse response : full.subList(0, full.size() - 1)) {
            check(response.hasChange(), "a full enumeration held something other than a change");
            check(!response.getChange().getUser().getId().isEmpty(), "a user change has no user id");
            if (!response.getChange().getRemoved()) {
                ids.add(response.getChange().getUser().getId());
            }
        }
        check(userId == null || ids.contains(userId), "a listed user is missing from the delta");

        List<DeltaUsersResponse> next = new ArrayList<>();
        stub().deltaUsers(DeltaUsersRequest.newBuilder().setDeltaLink(link).build())
                .forEachRemaining(next::add);
        check(!next.isEmpty() && !next.get(next.size() - 1).getDeltaLink().isEmpty(),
                "the second stream did not end with a delta link");
        check(next.stream().noneMatch(DeltaUsersResponse::hasResync), "the fresh link was refused");
        return ids.size() + " user(s); from the link, " + (next.size() - 1) + " change(s)";
    }

    private String membersOfAGroup() {
        need(groupId, "a group");
        List<DirectoryMember> members = new ArrayList<>();
        stub().listGroupTransitiveMembers(ListGroupTransitiveMembersRequest.newBuilder()
                .setGroupId(groupId).build())
                .forEachRemaining(response -> members.add(response.getMember()));
        return members.size() + " member(s), kinds " + memberKinds(members);
    }

    /** Reads of the directory a run waits through before it says a change never showed. */
    private static final int DIRECTORY_READS = 60;

    private String entraGroupName() {
        return listName + "-entra";
    }

    /** Waits between two reads of the directory. */
    private void pause() throws InterruptedException {
        Thread.sleep(backend.copyPollInterval() == null ? Duration.ofSeconds(2)
                : backend.copyPollInterval());
    }

    /** A skip when Graph refuses directory writes to this app; any other error as it is. */
    private static RuntimeException directoryWriteNotGranted(StatusRuntimeException e) {
        if (e.getStatus().getCode() != io.grpc.Status.Code.PERMISSION_DENIED) {
            return e;
        }
        return new Skip("the app may not write to the directory (Group.ReadWrite.All): "
                + e.getStatus().getDescription());
    }

    private String createGroup() {
        String name = entraGroupName();
        GraphGroup created;
        try {
            created = stub().createGroup(ai.pipestream.microsoft.v1.CreateGroupRequest.newBuilder()
                    .setGroup(GraphGroup.newBuilder().setDisplayName(name).setMailNickname(name)
                            .setDescription(name + " created by a live test run; safe to delete"))
                    .build()).getGroup();
        } catch (StatusRuntimeException e) {
            throw directoryWriteNotGranted(e);
        }
        check(!created.getId().isEmpty(), "the new group has no id");
        entraGroupId = created.getId();
        check(created.getDisplayName().equals(name), "the new group is called " + created.getDisplayName());
        check(created.getMailNickname().equals(name), "the new group's alias is " + created.getMailNickname());
        check(created.hasSecurityEnabled() && created.getSecurityEnabled(), "not a security group");
        check(created.hasMailEnabled() && !created.getMailEnabled(), "the group is mail-enabled");
        check(created.getGroupTypesCount() == 0, "group types " + created.getGroupTypesList());
        return "security group " + name + " created";
    }

    /** Reads the run's group until {@code seen} holds; null when the group is not found. */
    private GraphGroup awaitGroup(java.util.function.Predicate<GraphGroup> seen, int[] reads)
            throws InterruptedException {
        for (reads[0] = 1; reads[0] <= DIRECTORY_READS; reads[0]++) {
            try {
                GraphGroup group = stub().getGroup(ai.pipestream.microsoft.v1.GetGroupRequest.newBuilder()
                        .setGroupId(entraGroupId).build()).getGroup();
                if (seen.test(group)) {
                    return group;
                }
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != io.grpc.Status.Code.NOT_FOUND) {
                    throw e;
                }
                if (seen.test(null)) {
                    return null;
                }
            }
            pause();
        }
        throw new AssertionError("not visible after " + DIRECTORY_READS + " reads");
    }

    /**
     * Sends a write to the run's group again while Entra answers that the group, or
     * its member, is not there; answers the try that worked.
     */
    private int onceKnown(Runnable write) throws InterruptedException {
        for (int attempt = 1;; attempt++) {
            try {
                write.run();
                return attempt;
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != io.grpc.Status.Code.NOT_FOUND || attempt == DIRECTORY_READS) {
                    throw e;
                }
                pause();
            }
        }
    }

    private String getGroup() throws InterruptedException {
        need(entraGroupId, "the Entra group this run created");
        int[] reads = new int[1];
        GraphGroup group = awaitGroup(java.util.Objects::nonNull, reads);
        check(group.getId().equals(entraGroupId), "GetGroup answered another group");
        check(group.getDisplayName().equals(entraGroupName()), "the group is called " + group.getDisplayName());
        check(!group.getSecurityIdentifier().isEmpty(), "the group has no security identifier");
        return "read back on read " + reads[0];
    }

    private String updateGroup() throws InterruptedException {
        need(entraGroupId, "the Entra group this run created");
        String renamed = entraGroupName() + "-renamed";
        String description = entraGroupName() + " renamed by a live test run";
        int tries = onceKnown(() -> stub().updateGroup(ai.pipestream.microsoft.v1.UpdateGroupRequest.newBuilder()
                .setGroupId(entraGroupId).setDisplayName(renamed).setDescription(description).build()));
        int[] reads = new int[1];
        GraphGroup group = awaitGroup(g -> g != null && g.getDisplayName().equals(renamed), reads);
        check(group.getDescription().equals(description), "the description is " + group.getDescription());
        check(group.getMailNickname().equals(entraGroupName()), "the alias changed to "
                + group.getMailNickname());
        return "renamed and described on try " + tries + "; GetGroup showed it on read " + reads[0];
    }

    /** Reads the run's group's members until the user is, or is no longer, one of them. */
    private int awaitMembership(boolean member) throws InterruptedException {
        for (int read = 1; read <= DIRECTORY_READS; read++) {
            List<DirectoryMember> members = new ArrayList<>();
            try {
                stub().listGroupTransitiveMembers(ListGroupTransitiveMembersRequest.newBuilder()
                        .setGroupId(entraGroupId).build())
                        .forEachRemaining(response -> members.add(response.getMember()));
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != io.grpc.Status.Code.NOT_FOUND) {
                    throw e;
                }
            }
            if (members.stream().anyMatch(one -> one.getUser().getId().equals(userId)) == member) {
                return read;
            }
            pause();
        }
        throw new AssertionError("not visible after " + DIRECTORY_READS + " reads");
    }

    private String addGroupMember() throws InterruptedException {
        need(entraGroupId, "the Entra group this run created");
        need(userId, "a directory user");
        int tries = onceKnown(() -> stub().addGroupMember(
                ai.pipestream.microsoft.v1.AddGroupMemberRequest.newBuilder()
                        .setGroupId(entraGroupId).setMemberId(userId).build()));
        return "user added on try " + tries + "; ListGroupTransitiveMembers showed it on read "
                + awaitMembership(true);
    }

    private String removeGroupMember() throws InterruptedException {
        need(entraGroupId, "the Entra group this run created");
        need(userId, "a directory user");
        int tries = onceKnown(() -> stub().removeGroupMember(
                ai.pipestream.microsoft.v1.RemoveGroupMemberRequest.newBuilder()
                        .setGroupId(entraGroupId).setMemberId(userId).build()));
        return "user removed on try " + tries + "; ListGroupTransitiveMembers showed it on read "
                + awaitMembership(false);
    }

    private String deleteGroup() throws InterruptedException {
        need(entraGroupId, "the Entra group this run created");
        int tries = onceKnown(() -> stub().deleteGroup(ai.pipestream.microsoft.v1.DeleteGroupRequest.newBuilder()
                .setGroupId(entraGroupId).build()));
        int[] reads = new int[1];
        awaitGroup(java.util.Objects::isNull, reads);
        entraGroupId = null;
        return "deleted on try " + tries + "; GetGroup answered NOT_FOUND on read " + reads[0];
    }

    private String ownersOfAGroup() {
        need(groupId, "a group");
        List<DirectoryMember> owners = new ArrayList<>();
        stub().listGroupOwners(ai.pipestream.microsoft.v1.ListGroupOwnersRequest.newBuilder()
                .setGroupId(groupId).build())
                .forEachRemaining(response -> owners.add(response.getOwner()));
        List<com.google.protobuf.Message> groupAndOwners = new ArrayList<>();
        groupAndOwners.add(listedGroup);
        groupAndOwners.addAll(owners);
        records.put("entra-group-owners", groupAndOwners);
        return owners.size() + " owner(s), kinds " + memberKinds(owners);
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

    /** The test folder as a list item of its library. */
    private SiteSecurable folderSecurable() {
        if (folderSecurable == null) {
            String library = stub().getDrive(ai.pipestream.microsoft.v1.GetDriveRequest.newBuilder()
                    .setDriveId(driveId).build()).getDrive().getListId();
            String number = stub().getItem(GetItemRequest.newBuilder().setDriveId(driveId)
                    .setItemId(need(folderId, "the test folder")).build()).getItem()
                    .getSharepointIds().getListItemId();
            check(!library.isEmpty() && !number.isEmpty(), "the folder's list or list item id is unknown");
            folderSecurable = SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder()
                    .setSiteId(siteId).setListId(library).setItemId(number)).build();
        }
        return folderSecurable;
    }

    private String breakInheritance() {
        need(itemId, "the list item");
        SiteSecurable item = SiteSecurable.newBuilder().setListItem(listItemRef(itemId)).build();
        for (SiteSecurable securable : List.of(item, folderSecurable())) {
            boolean unique = stub().setSiteRoleInheritance(
                    ai.pipestream.microsoft.v1.SetSiteRoleInheritanceRequest.newBuilder()
                            .setSecurable(securable).setCopyRoleAssignments(true).build())
                    .getHasUniqueRoleAssignments();
            check(unique, "still inheriting after the break");
            check(stub().getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                    .setSecurable(securable).build()).getHasUniqueRoleAssignments(),
                    "GetSiteRoleInheritance still reports inheritance");
        }
        inheritanceBroken = true;
        return "the list item and the test folder have their own role assignments, copied from the parent";
    }

    private String syncSecurity() {
        if (!inheritanceBroken) {
            throw new Skip("needs the list item and folder with their own role assignments");
        }
        need(small, "the small file");
        Map<String, ai.pipestream.microsoft.v1.MicrosoftEntity> entities = new java.util.LinkedHashMap<>();
        try {
            blocking.withDeadlineAfter(5, TimeUnit.MINUTES).sync(
                    ai.pipestream.microsoft.v1.SyncRequest.newBuilder().addSiteIds(siteId)
                            .setIncludeLists(true)
                            .setInclude(ai.pipestream.microsoft.v1.DriveItemInclude.getDefaultInstance())
                            .setIncludeSecurity(true).build())
                    .forEachRemaining(response -> {
                        if (response.hasChange()) {
                            ai.pipestream.microsoft.v1.MicrosoftEntity entity = response.getChange().getEntity();
                            entities.put(entity.getEntityCase() + " " + entity.getEntityId(), entity);
                        } else if (!response.getResumeCursor().isEmpty()) {
                            securityCursor = response.getResumeCursor();
                        }
                    });
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        check(securityCursor != null, "the pass ended without a resume cursor");
        records.put("sharepoint-sync-records", List.copyOf(entities.values()));
        String siteAcl = MicrosoftEntityIds.siteAcl(SiteSecurable.newBuilder().setSiteId(siteId).build());
        String itemAcl = MicrosoftEntityIds.siteAcl(
                SiteSecurable.newBuilder().setListItem(listItemRef(itemId)).build());
        String folderAcl = MicrosoftEntityIds.siteAcl(folderSecurable());
        int groups = 0;
        int acls = 0;
        int owned = 0;
        int users = 0;
        int joinable = 0;
        int admins = 0;
        for (ai.pipestream.microsoft.v1.MicrosoftEntity entity : entities.values()) {
            String owner = switch (entity.getEntityCase()) {
                case DRIVE -> entity.getDrive().getListId().isEmpty() ? null : entity.getDrive().getAclOwnerId();
                case DRIVE_ITEM -> entity.getDriveItem().getAclOwnerId();
                case LIST -> entity.getList().getAclOwnerId();
                case LIST_ITEM -> entity.getListItem().getAclOwnerId();
                default -> null;
            };
            if (owner != null) {
                check(entities.containsKey("SITE_ACL " + owner),
                        "an entity names role assignments that were not emitted");
                owned++;
            }
            if (entity.hasSiteGroup()) {
                groups++;
                check(entity.getSiteGroup().getCompleteness()
                        == ai.pipestream.sync.v1.AclCompleteness.ACL_COMPLETENESS_COMPLETE,
                        "a group's members are not complete");
                entity.getSiteGroup().getMembersList().forEach(WriteScenarios::checkPrincipal);
            }
            if (entity.hasSiteAdmins()) {
                admins += entity.getSiteAdmins().getAdminsCount();
                check(entity.getSiteAdmins().getCompleteness()
                        == ai.pipestream.sync.v1.AclCompleteness.ACL_COMPLETENESS_COMPLETE,
                        "the site's administrators are not complete");
                for (ai.pipestream.microsoft.v1.SitePrincipal admin : entity.getSiteAdmins().getAdminsList()) {
                    checkPrincipal(admin);
                    check(admin.getSiteAdmin(), "a listed administrator is not marked as one");
                }
            }
            if (entity.hasSiteAcl()) {
                acls++;
                check(entity.getSiteAcl().getCompleteness()
                        == ai.pipestream.sync.v1.AclCompleteness.ACL_COMPLETENESS_COMPLETE,
                        "role assignments are not complete");
                check(entity.getSiteAcl().getAssignmentsCount() > 0, "role assignments are empty");
                for (SiteRoleAssignment assignment : entity.getSiteAcl().getAssignmentsList()) {
                    checkPrincipal(assignment.getMember());
                    if (assignment.getMember().getPrincipalType()
                            == SitePrincipalType.SITE_PRINCIPAL_TYPE_USER) {
                        users++;
                        joinable += assignment.getMember().getUserPrincipalName().isEmpty() ? 0 : 1;
                    }
                }
            }
        }
        check(groups > 0, "no site group was emitted");
        check(entities.containsKey("SITE_ADMINS " + MicrosoftEntityIds.siteAdmins(siteId)),
                "the site's administrators were not emitted");
        check(admins > 0, "the site has no administrator");
        for (String acl : List.of(siteAcl, itemAcl, folderAcl)) {
            check(entities.containsKey("SITE_ACL " + acl), "expected role assignments were not emitted");
        }
        ai.pipestream.microsoft.v1.MicrosoftEntity list = entities.get("LIST " + listId);
        check(list != null && list.getList().getAclOwnerId().equals(siteAcl),
                "the test list does not point at the site's role assignments");
        ai.pipestream.microsoft.v1.MicrosoftEntity item = entities.get("LIST_ITEM " + listId + "/" + itemId);
        check(item != null && item.getListItem().getAclOwnerId().equals(itemAcl),
                "the list item does not point at its own role assignments");
        ai.pipestream.microsoft.v1.MicrosoftEntity folder = entities.get("DRIVE_ITEM " + folderId);
        check(folder != null && folder.getDriveItem().getAclOwnerId().equals(folderAcl),
                "the test folder does not point at its own role assignments");
        ai.pipestream.microsoft.v1.MicrosoftEntity file = entities.get("DRIVE_ITEM " + small.getId());
        check(file != null && file.getDriveItem().getAclOwnerId().equals(folderAcl),
                "a file in the test folder does not point at the folder's role assignments");
        String sharing = "";
        if (sharedSecurable != null) {
            String sharedAcl = MicrosoftEntityIds.siteAcl(sharedSecurable);
            ai.pipestream.microsoft.v1.MicrosoftEntity sharedFile = entities.get("DRIVE_ITEM " + shared.getId());
            check(sharedFile != null && sharedFile.getDriveItem().getAclOwnerId().equals(sharedAcl),
                    "the shared file does not point at its own role assignments");
            ai.pipestream.microsoft.v1.MicrosoftEntity acl = entities.get("SITE_ACL " + sharedAcl);
            check(acl != null, "the role assignments of the shared file were not emitted");
            Set<String> scopes = new TreeSet<>();
            for (SiteRoleAssignment assignment : acl.getSiteAcl().getAssignmentsList()) {
                if (assignment.getMember().hasSharingLink()) {
                    scopes.add(assignment.getMember().getSharingLink().getScope());
                }
            }
            check(scopes.contains("organization"),
                    "the shared file's role assignments name no group of an organization link: " + scopes);
            long linkGroups = entities.values().stream().filter(entity -> entity.hasSiteGroup()
                    && entity.getSiteGroup().getGroup().hasSharingLink()).count();
            check(linkGroups >= scopes.size(), "a link group was not emitted with its members");
            sharing = "; the shared file has its own role assignments, with link groups for " + scopes
                    + ", and " + linkGroups + " link group(s) arrive with their members";
        }
        return groups + " group(s), " + admins + " site administrator(s), " + acls
                + " role assignment set(s), " + owned
                + " drives, lists and items with an owner, " + joinable + " of " + users
                + " user grantee(s) with a user principal name" + sharing;
    }

    private String createSiteGroup() {
        String title = listName + "-group";
        SiteGroup group;
        try {
            group = stub().createSiteGroup(ai.pipestream.microsoft.v1.CreateSiteGroupRequest.newBuilder().setSiteId(siteId)
                    .setTitle(title).setDescription("Created by a live test run; safe to delete.").build())
                    .getGroup();
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        check(group.getId() > 0, "the new group has no id");
        ownGroupId = group.getId();
        createdGroupId = group.getId();
        check(group.getTitle().equals(title), "the new group is called " + group.getTitle());
        check(siteGroupIds().contains(group.getId()), "ListSiteGroups does not show the new group");
        return "group " + title + " created; ListSiteGroups shows it";
    }

    private Set<Integer> siteGroupIds() {
        Set<Integer> ids = new TreeSet<>();
        stub().listSiteGroups(ListSiteGroupsRequest.newBuilder().setSiteId(siteId).build())
                .forEachRemaining(response -> ids.add(response.getGroup().getId()));
        return ids;
    }

    private String roleDefinitions() {
        List<ai.pipestream.microsoft.v1.SiteRoleDefinition> roles = new ArrayList<>();
        try {
            stub().listSiteRoleDefinitions(ai.pipestream.microsoft.v1.ListSiteRoleDefinitionsRequest.newBuilder()
                    .setSiteId(siteId).build()).forEachRemaining(response -> roles.add(response.getRole()));
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        check(roles.stream().allMatch(role -> role.getId() > 0 && !role.getName().isEmpty()),
                "a permission level has no id or no name");
        check(roles.stream().map(ai.pipestream.microsoft.v1.SiteRoleDefinition::getId).distinct().count() == roles.size(),
                "a permission level was listed twice");
        // SharePoint's role type 2 is the reader.
        readRoleId = roles.stream().filter(role -> role.getRoleTypeKind() == 2)
                .map(ai.pipestream.microsoft.v1.SiteRoleDefinition::getId).findFirst().orElse(null);
        check(readRoleId != null, "the site has no read permission level");
        return roles.size() + " permission level(s): "
                + roles.stream().map(ai.pipestream.microsoft.v1.SiteRoleDefinition::getName).sorted().toList();
    }

    private List<SitePrincipal> membersOfTheOwnGroup() {
        List<SitePrincipal> members = new ArrayList<>();
        stub().listSiteGroupMembers(ListSiteGroupMembersRequest.newBuilder().setSiteId(siteId)
                .setGroupId(ownGroupId).build())
                .forEachRemaining(response -> members.add(response.getMember()));
        return members;
    }

    private String addSiteGroupMember() {
        need(ownGroupId, "the group this run created");
        need(userPrincipalName, "a directory user with a sign-in name");
        SitePrincipal member = stub().addSiteGroupMember(ai.pipestream.microsoft.v1.AddSiteGroupMemberRequest.newBuilder()
                .setSiteId(siteId).setGroupId(ownGroupId).setUserPrincipalName(userPrincipalName).build())
                .getMember();
        check(member.getId() > 0, "the new member has no id");
        ownGroupMemberId = member.getId();
        checkPrincipal(member);
        check(member.getPrincipalType() == SitePrincipalType.SITE_PRINCIPAL_TYPE_USER,
                "the new member is a " + member.getPrincipalType());
        check(member.getUserPrincipalName().equalsIgnoreCase(userPrincipalName),
                "the new member has another user principal name");
        List<SitePrincipal> members = membersOfTheOwnGroup();
        check(members.size() == 1 && members.get(0).getId() == member.getId(),
                "ListSiteGroupMembers shows " + members.size() + " member(s), not the one added");
        return "the user is a member; ListSiteGroupMembers shows it";
    }

    /** The permission levels the run's own group holds on the list item. */
    private List<Long> rolesOfTheOwnGroupOnTheItem() {
        List<Long> roles = new ArrayList<>();
        stub().listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(listItemRef(itemId))).build())
                .forEachRemaining(response -> {
                    if (response.getAssignment().getPrincipalId() == ownGroupId) {
                        response.getAssignment().getRolesList().forEach(role -> roles.add(role.getId()));
                    }
                });
        return roles;
    }

    private String addRoleAssignment() {
        need(ownGroupId, "the group this run created");
        need(readRoleId, "the read permission level");
        if (!inheritanceBroken) {
            throw new Skip("needs the list item with its own role assignments");
        }
        stub().addSiteRoleAssignment(ai.pipestream.microsoft.v1.AddSiteRoleAssignmentRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(listItemRef(itemId)))
                .setPrincipalId(ownGroupId).setRoleDefinitionId(readRoleId).build());
        check(rolesOfTheOwnGroupOnTheItem().equals(List.of(readRoleId)),
                "ListSiteRoleAssignments does not show the group with the read level alone");
        return "the group holds the read level on the list item; ListSiteRoleAssignments shows it";
    }

    private String removeRoleAssignment() {
        need(ownGroupId, "the group this run created");
        need(readRoleId, "the read permission level");
        if (!inheritanceBroken) {
            throw new Skip("needs the list item with its own role assignments");
        }
        stub().removeSiteRoleAssignment(ai.pipestream.microsoft.v1.RemoveSiteRoleAssignmentRequest.newBuilder()
                .setSecurable(SiteSecurable.newBuilder().setListItem(listItemRef(itemId)))
                .setPrincipalId(ownGroupId).setRoleDefinitionId(readRoleId).build());
        check(rolesOfTheOwnGroupOnTheItem().isEmpty(),
                "ListSiteRoleAssignments still shows the group on the list item");
        return "the group holds nothing on the list item any more";
    }

    private String removeSiteGroupMember() {
        need(ownGroupId, "the group this run created");
        need(ownGroupMemberId, "the member this run added");
        stub().removeSiteGroupMember(ai.pipestream.microsoft.v1.RemoveSiteGroupMemberRequest.newBuilder().setSiteId(siteId)
                .setGroupId(ownGroupId).setMemberId(ownGroupMemberId).build());
        check(membersOfTheOwnGroup().isEmpty(), "ListSiteGroupMembers still shows a member");
        ownGroupMemberId = null;
        return "the group has no member any more";
    }

    private String deleteSiteGroup() {
        need(ownGroupId, "the group this run created");
        int id = ownGroupId;
        stub().deleteSiteGroup(ai.pipestream.microsoft.v1.DeleteSiteGroupRequest.newBuilder().setSiteId(siteId)
                .setGroupId(id).build());
        ownGroupId = null;
        check(!siteGroupIds().contains(id), "ListSiteGroups still shows the group");
        return "the group is gone from ListSiteGroups";
    }

    private String restoreInheritance() {
        if (!inheritanceBroken) {
            throw new Skip("nothing stopped inheriting");
        }
        SiteSecurable item = SiteSecurable.newBuilder().setListItem(listItemRef(itemId)).build();
        for (SiteSecurable securable : List.of(item, folderSecurable())) {
            boolean unique = stub().setSiteRoleInheritance(
                    ai.pipestream.microsoft.v1.SetSiteRoleInheritanceRequest.newBuilder()
                            .setSecurable(securable).setInherit(true).build())
                    .getHasUniqueRoleAssignments();
            check(!unique, "still has its own role assignments after the reset");
        }
        inheritanceBroken = false;
        return "the list item and the test folder inherit again";
    }

    private String syncSecurityResumed() {
        if (securityCursor == null || inheritanceBroken) {
            throw new Skip("needs a security sync from before the reset");
        }
        Map<String, ai.pipestream.microsoft.v1.MicrosoftChange> changes = new java.util.LinkedHashMap<>();
        blocking.withDeadlineAfter(5, TimeUnit.MINUTES).sync(
                ai.pipestream.microsoft.v1.SyncRequest.newBuilder().addSiteIds(siteId)
                        .setIncludeLists(true)
                        .setInclude(ai.pipestream.microsoft.v1.DriveItemInclude.getDefaultInstance())
                        .setIncludeSecurity(true).setSinceCursor(securityCursor).build())
                .forEachRemaining(response -> {
                    if (response.hasChange()) {
                        ai.pipestream.microsoft.v1.MicrosoftEntity entity = response.getChange().getEntity();
                        changes.put(entity.getEntityCase() + " " + entity.getEntityId(), response.getChange());
                    }
                });
        String itemAcl = MicrosoftEntityIds.siteAcl(
                SiteSecurable.newBuilder().setListItem(listItemRef(itemId)).build());
        String folderAcl = MicrosoftEntityIds.siteAcl(folderSecurable());
        ai.pipestream.microsoft.v1.MicrosoftChange list = changes.get("LIST " + listId);
        ai.pipestream.microsoft.v1.MicrosoftChange drive = changes.get("DRIVE " + driveId);
        check(list != null && drive != null, "the pass did not emit the test list and its library");
        String listOwner = list.getEntity().getList().getAclOwnerId();
        String driveOwner = drive.getEntity().getDrive().getAclOwnerId();
        check(!driveOwner.isEmpty() && !listOwner.isEmpty(), "the list or the library has no owner");
        ai.pipestream.microsoft.v1.MicrosoftChange item = changes.get("LIST_ITEM " + listId + "/" + itemId);
        check(item != null, "the list item that inherits again was not emitted");
        check(item.getEntity().getListItem().getAclOwnerId().equals(listOwner),
                "the list item still points at " + item.getEntity().getListItem().getAclOwnerId());
        ai.pipestream.microsoft.v1.MicrosoftChange folder = changes.get("DRIVE_ITEM " + folderId);
        check(folder != null, "the folder that inherits again was not emitted");
        check(folder.getEntity().getDriveItem().getAclOwnerId().equals(driveOwner),
                "the folder still points at " + folder.getEntity().getDriveItem().getAclOwnerId());
        ai.pipestream.microsoft.v1.MicrosoftChange file = changes.get("DRIVE_ITEM " + small.getId());
        check(file != null, "the file in the folder was not emitted");
        check(file.getEntity().getDriveItem().getAclOwnerId().equals(driveOwner),
                "the file in the folder still points at " + file.getEntity().getDriveItem().getAclOwnerId());
        for (String acl : List.of(itemAcl, folderAcl)) {
            ai.pipestream.microsoft.v1.MicrosoftChange gone = changes.get("SITE_ACL " + acl);
            check(gone != null && gone.getOperation()
                    == ai.pipestream.microsoft.v1.ChangeOperation.CHANGE_OPERATION_DELETE,
                    "role assignments that no longer exist were not deleted");
        }
        long items = changes.values().stream().filter(change -> change.getEntity().hasDriveItem()
                || change.getEntity().hasListItem()).count();
        return "the list item, the folder and the file in it point at their parent's role assignments again; "
                + items + " item(s) in the pass, the two role assignment sets that are gone arrive as deletes";
    }

    private static final List<SiteChangeScope> CHANGE_SCOPES = List.of(
            SiteChangeScope.SITE_CHANGE_SCOPE_WEB, SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION,
            SiteChangeScope.SITE_CHANGE_SCOPE_LIST);

    /** A ListSiteChanges request for one of the three logs of the test site. */
    private ListSiteChangesRequest.Builder changesOf(SiteChangeScope scope) {
        ListSiteChangesRequest.Builder request = ListSiteChangesRequest.newBuilder().setSiteId(siteId)
                .setScope(scope);
        return scope == SiteChangeScope.SITE_CHANGE_SCOPE_LIST ? request.setListId(listId) : request;
    }

    /** Reads a change log; the token the stream ended with is the last element's. */
    private List<ListSiteChangesResponse> changes(ListSiteChangesRequest request) {
        List<ListSiteChangesResponse> stream = new ArrayList<>();
        try {
            stub().listSiteChanges(request).forEachRemaining(stream::add);
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        check(!stream.isEmpty() && stream.get(stream.size() - 1).getEventCase()
                == ListSiteChangesResponse.EventCase.CHANGE_TOKEN
                && !stream.get(stream.size() - 1).getChangeToken().isEmpty(),
                "the stream did not end with a change token");
        return stream;
    }

    private String endOfTheChangeLogs() {
        need(listId, "the test list");
        for (SiteChangeScope scope : CHANGE_SCOPES) {
            List<ListSiteChangesResponse> stream = changes(changesOf(scope)
                    .setChangeToken("latest").build());
            check(stream.size() == 1, "the end of a log came with " + (stream.size() - 1) + " change(s)");
            changeTokens.put(scope, stream.get(0).getChangeToken());
        }
        check(new HashSet<>(changeTokens.values()).size() == CHANGE_SCOPES.size(),
                "two logs ended on the same token");
        return "a change token for the site, the site collection and the test list";
    }

    private String siteChanges() {
        if (changeTokens.size() < CHANGE_SCOPES.size() || inheritanceBroken) {
            throw new Skip("needs the change tokens from before the list item stopped inheriting, "
                    + "and the item inheriting again");
        }
        need(itemId, "the list item");
        List<String> seen = new ArrayList<>();
        List<String> groups = new ArrayList<>();
        for (SiteChangeScope scope : CHANGE_SCOPES) {
            List<ListSiteChangesResponse> stream = changes(changesOf(scope)
                    .setChangeToken(changeTokens.get(scope)).build());
            String end = stream.get(stream.size() - 1).getChangeToken();
            check(!end.equals(changeTokens.get(scope)), "the log did not move");
            Map<String, Integer> kinds = new TreeMap<>();
            Set<SiteChangeType> onTheItem = new TreeSet<>();
            Set<SiteChangeType> onTheGroup = new TreeSet<>();
            for (ListSiteChangesResponse response : stream.subList(0, stream.size() - 1)) {
                SiteChange change = response.getChange();
                check(change.hasTime() && !change.getChangeToken().isEmpty(),
                        "a change has no time or no token");
                check(change.getObject() != SiteChangeObject.SITE_CHANGE_OBJECT_UNSPECIFIED
                        && change.getType() != SiteChangeType.SITE_CHANGE_TYPE_UNSPECIFIED,
                        "a change of an unknown kind: " + change.getMetadataMap());
                kinds.merge(name(change.getObject()) + " " + name(change.getType()), 1, Integer::sum);
                if (change.getObject() == SiteChangeObject.SITE_CHANGE_OBJECT_ITEM
                        && change.getListId().equalsIgnoreCase(listId)
                        && change.getItemId().equals(itemId)) {
                    check(change.getSecurable().getListItem().getItemId().equals(itemId),
                            "the change on the list item does not name it as a securable");
                    onTheItem.add(change.getType());
                }
                if (change.getObject() == SiteChangeObject.SITE_CHANGE_OBJECT_GROUP
                        && createdGroupId != null && change.getGroupId() == createdGroupId) {
                    onTheGroup.add(change.getType());
                }
            }
            // Groups belong to the site collection; only its log reports them.
            if (scope == SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION && createdGroupId != null
                    && ownGroupId == null) {
                check(onTheGroup.containsAll(List.of(SiteChangeType.SITE_CHANGE_TYPE_ADD,
                        SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_ADD,
                        SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_DELETE,
                        SiteChangeType.SITE_CHANGE_TYPE_DELETE)),
                        "the site collection log does not show the run's group created, given and "
                        + "losing a member, and deleted: " + onTheGroup + " of " + kinds);
            }
            check(onTheItem.contains(SiteChangeType.SITE_CHANGE_TYPE_SCOPE_ADD)
                    && onTheItem.contains(SiteChangeType.SITE_CHANGE_TYPE_SCOPE_DELETE),
                    "the " + name(scope) + " log does not show the list item stop inheriting "
                    + "and inherit again: " + onTheItem + " of " + kinds);
            // Read again from the end: nothing new, same token.
            List<ListSiteChangesResponse> again = changes(changesOf(scope).setChangeToken(end)
                    .addObjects(SiteChangeObject.SITE_CHANGE_OBJECT_ITEM).build());
            check(again.stream().filter(ListSiteChangesResponse::hasChange).allMatch(response ->
                    response.getChange().getObject() == SiteChangeObject.SITE_CHANGE_OBJECT_ITEM),
                    "a filter for items answered another object");
            seen.add(name(scope) + ": " + kinds + ", on the list item "
                    + onTheItem.stream().map(WriteScenarios::name).toList());
            Map<String, Integer> ofGroups = new TreeMap<>(kinds);
            ofGroups.keySet().removeIf(kind -> !kind.startsWith("GROUP ") && !kind.startsWith("USER "));
            groups.add(name(scope) + " " + ofGroups);
        }
        return "groups and users: " + String.join(", ", groups) + "; " + String.join("; ", seen);
    }

    private String siteChangesFromAnOldToken() {
        String token = changeTokens.get(SiteChangeScope.SITE_CHANGE_SCOPE_LIST);
        need(token, "a change token of the test list");
        // version;scope;id;ticks;number - the same log, a year earlier.
        String[] parts = token.split(";");
        check(parts.length == 5, "a change token of " + parts.length + " parts");
        long ticks = Long.parseLong(parts[3]) - Duration.ofDays(365).toSeconds() * 10_000_000L;
        String old = String.join(";", parts[0], parts[1], parts[2], Long.toString(ticks), "1");
        List<ListSiteChangesResponse> stream = changes(changesOf(SiteChangeScope.SITE_CHANGE_SCOPE_LIST)
                .setChangeToken(old).build());
        check(stream.size() > 1, "a token older than the list answered no change");
        return "SharePoint accepted the token: " + (stream.size() - 1) + " change(s)";
    }

    private String siteChangesFromABadToken() {
        try {
            stub().listSiteChanges(changesOf(SiteChangeScope.SITE_CHANGE_SCOPE_WEB)
                    .setChangeToken("not-a-change-token").build()).forEachRemaining(response -> { });
        } catch (StatusRuntimeException e) {
            check(e.getStatus().getCode() == io.grpc.Status.Code.INVALID_ARGUMENT,
                    "refused with " + e.getStatus().getCode() + ": " + e.getStatus().getDescription());
            return "refused as INVALID_ARGUMENT: " + e.getStatus().getDescription();
        }
        throw new AssertionError("SharePoint accepted a token that is not one");
    }

    private static String name(com.google.protobuf.ProtocolMessageEnum value) {
        String name = value.getValueDescriptor().getName();
        return name.substring(name.indexOf("_", "SITE_CHANGE_".length()) + 1);
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

    // ------------------------------------------------------------ sharing

    private PermissionTarget sharedTarget() {
        return PermissionTarget.newBuilder().setDriveItem(ref(need(shared, "the shared file").getId())).build();
    }

    private Permission permissionOfTheSharedFile(String id) {
        return stub().getPermission(GetPermissionRequest.newBuilder().setTarget(sharedTarget())
                .setPermissionId(id).build()).getPermission();
    }

    private SharingRecipient theUser() {
        return SharingRecipient.newBuilder().setObjectId(need(userId, "a directory user")).build();
    }

    /** Whether a permission names the user or group with this directory id as a grantee. */
    private static boolean names(Permission permission, String id) {
        List<SharePointIdentitySet> grantees = new ArrayList<>(permission.getGrantedToIdentitiesList());
        grantees.add(permission.getGrantedTo());
        return grantees.stream().anyMatch(grantee -> grantee.getUser().getId().equalsIgnoreCase(id)
                || grantee.getGroup().getId().equalsIgnoreCase(id));
    }

    private String shareWithTheOrganization() throws Exception {
        shared = upload("shared.txt",
                remember(("shared " + folderName + "\n").getBytes(StandardCharsets.UTF_8)), 0,
                ConflictBehavior.CONFLICT_BEHAVIOR_FAIL);
        CreateSharingLinkRequest request = CreateSharingLinkRequest.newBuilder().setTarget(sharedTarget())
                .setType("view").setScope("organization").build();
        CreateSharingLinkResponse made = stub().createSharingLink(request);
        Permission link = made.getPermission();
        check(made.getCreated(), "the first link of the file was not reported as new");
        check(!link.getId().isEmpty(), "the link has no id");
        check(link.getLink().getScope().equals("organization") && link.getLink().getType().equals("view"),
                "asked for a view link for the organization, got " + link.getLink().getType() + " for "
                        + link.getLink().getScope());
        check(!link.getLink().getWebUrl().isEmpty(), "the link has no address");
        check(link.getTarget().equals(sharedTarget()), "the link names another target");
        CreateSharingLinkResponse again = stub().createSharingLink(request);
        check(!again.getCreated() && again.getPermission().getId().equals(link.getId()),
                "asking for the same link again made another one");
        organizationLinkId = link.getId();
        return "a view link for the organization with roles " + link.getRolesList()
                + "; asking again answers the same link and says it is not new";
    }

    private String getPermission() {
        need(organizationLinkId, "the organization link");
        Permission read = permissionOfTheSharedFile(organizationLinkId);
        check(read.getId().equals(organizationLinkId), "another permission was read");
        check(read.getLink().getScope().equals("organization"), "the link read back is for "
                + read.getLink().getScope());
        check(read.getTarget().equals(sharedTarget()), "the permission names another target");
        io.grpc.Status.Code missing;
        try {
            permissionOfTheSharedFile("pipestream-no-such-permission");
            throw new AssertionError("a permission that does not exist was read");
        } catch (StatusRuntimeException e) {
            missing = e.getStatus().getCode();
        }
        check(missing == io.grpc.Status.Code.NOT_FOUND || missing == io.grpc.Status.Code.INVALID_ARGUMENT,
                "a permission that does not exist answers " + missing);
        return "the link read back by id; an id that does not exist answers " + missing;
    }

    private String shareWithNamedPeople() {
        need(organizationLinkId, "the shared file");
        Permission link = stub().createSharingLink(CreateSharingLinkRequest.newBuilder()
                .setTarget(sharedTarget()).setType("view").setScope("users").build()).getPermission();
        check(link.getLink().getScope().equals("users"), "asked for a link for named people, got one for "
                + link.getLink().getScope());
        check(!link.getId().isEmpty() && !link.getId().equals(organizationLinkId),
                "the link for named people is the organization link");
        usersLinkId = link.getId();
        return "a view link for named people, " + link.getGrantedToIdentitiesCount() + " of them so far";
    }

    private String grantLinkAccess() {
        need(usersLinkId, "the link for named people");
        List<Permission> granted = stub().grantLinkAccess(GrantLinkAccessRequest.newBuilder()
                .setItem(ref(shared.getId())).setPermissionId(usersLinkId).addRecipients(theUser())
                .addRoles("read").build()).getPermissionsList();
        check(!granted.isEmpty(), "the grant answered no permission");
        check(granted.stream().allMatch(permission -> permission.getTarget().equals(sharedTarget())),
                "a granted permission names another target");
        Permission link = permissionOfTheSharedFile(usersLinkId);
        check(names(link, userId), "the link does not name the user it was granted to");
        return granted.size() + " permission(s) answered; the link read back names the user among "
                + link.getGrantedToIdentitiesCount() + " grantee(s)";
    }

    private String inviteRecipients() {
        need(organizationLinkId, "the shared file");
        need(groupId, "a directory group");
        stub().listPermissions(ListPermissionsRequest.newBuilder().setTarget(sharedTarget()).build())
                .forEachRemaining(response -> {
                    if (!response.getPermission().hasLink() && names(response.getPermission(), groupId)) {
                        rolesBeforeTheInvitation.addAll(response.getPermission().getRolesList());
                    }
                });
        check(rolesBeforeTheInvitation.isEmpty(), "the group already holds " + rolesBeforeTheInvitation
                + " on the file");
        List<Permission> made = stub().inviteRecipients(InviteRecipientsRequest.newBuilder()
                .setItem(ref(shared.getId()))
                .addRecipients(SharingRecipient.newBuilder().setObjectId(groupId)).addRoles("read")
                .setRequireSignIn(true).build()).getPermissionsList();
        check(made.size() == 1, "one recipient got " + made.size() + " permission(s)");
        Permission permission = made.get(0);
        check(!permission.getId().isEmpty(), "the permission has no id");
        check(permission.getRolesList().contains("read"), "asked for read, got " + permission.getRolesList());
        check(permission.getTarget().equals(sharedTarget()), "the permission names another target");
        invitedPermissionId = permission.getId();
        return "1 permission with roles " + permission.getRolesList() + " for a "
                + (permission.hasGrantedTo() ? kind(permission.getGrantedTo()) : "grantee Graph does not name")
                + ", a link " + (permission.hasLink() ? "for " + permission.getLink().getScope()
                        : "is not part of it") + ", no mail asked for";
    }

    private String updatePermission() {
        need(invitedPermissionId, "the invited group's permission");
        Permission changed = stub().updatePermission(UpdatePermissionRequest.newBuilder()
                .setTarget(sharedTarget()).setPermissionId(invitedPermissionId).addRoles("write").build())
                .getPermission();
        check(changed.getRolesList().contains("write"), "asked for write, got " + changed.getRolesList());
        check(!changed.getId().isEmpty(), "the changed permission has no id");
        String before = invitedPermissionId;
        invitedPermissionId = changed.getId();
        Permission read = permissionOfTheSharedFile(invitedPermissionId);
        check(read.getRolesList().contains("write"), "read back with roles " + read.getRolesList());
        return "roles " + read.getRolesList() + " read back, under "
                + (before.equals(invitedPermissionId) ? "the same id" : "a new id");
    }

    private String shareWithAnyone() {
        need(organizationLinkId, "the shared file");
        Permission link;
        try {
            link = stub().createSharingLink(CreateSharingLinkRequest.newBuilder().setTarget(sharedTarget())
                    .setType("view").setScope("anonymous").build()).getPermission();
        } catch (StatusRuntimeException e) {
            io.grpc.Status.Code code = e.getStatus().getCode();
            if (code != io.grpc.Status.Code.PERMISSION_DENIED && code != io.grpc.Status.Code.INVALID_ARGUMENT
                    && code != io.grpc.Status.Code.FAILED_PRECONDITION) {
                throw e;
            }
            return "the site refuses a link for anyone: " + code;
        }
        check(!link.getId().isEmpty(), "the link has no id");
        if (!link.getId().equals(organizationLinkId) && !link.getId().equals(usersLinkId)) {
            anyoneLinkId = link.getId();
        }
        check(!link.getLink().getScope().isEmpty(), "the link has no scope");
        return link.getLink().getScope().equals("anonymous")
                ? "the site allows it: a view link for anyone with roles " + link.getRolesList()
                : "asked for a link for anyone, the site answered one for " + link.getLink().getScope();
    }

    private String permissionsOfTheSharedFile() {
        need(organizationLinkId, "the organization link");
        Map<String, Permission> byId = new HashMap<>();
        Map<String, Integer> scopes = new TreeMap<>();
        stub().listPermissions(ListPermissionsRequest.newBuilder().setTarget(sharedTarget()).build())
                .forEachRemaining(response -> {
                    Permission permission = response.getPermission();
                    byId.put(permission.getId(), permission);
                    if (permission.hasLink()) {
                        scopes.merge(permission.getLink().getScope(), 1, Integer::sum);
                    }
                });
        Permission organization = byId.get(organizationLinkId);
        check(organization != null && organization.getLink().getScope().equals("organization"),
                "the ACL does not list the organization link");
        if (usersLinkId != null) {
            Permission people = byId.get(usersLinkId);
            check(people != null && people.getLink().getScope().equals("users"),
                    "the ACL does not list the link for named people");
            check(userId == null || names(people, userId),
                    "the link for named people does not name the user it was granted to");
        }
        if (anyoneLinkId != null) {
            check(byId.containsKey(anyoneLinkId), "the ACL does not list the link for anyone");
        }
        if (invitedPermissionId != null) {
            Permission invited = byId.get(invitedPermissionId);
            check(invited != null && invited.getRolesList().contains("write"),
                    "the ACL does not list the invited group with write"
                            + (invited == null ? "" : ", but with " + invited.getRolesList()));
        }
        return byId.size() + " permission(s), links by scope " + scopes
                + (invitedPermissionId == null ? "" : ", the invited group with write");
    }

    /** A drive item of the test library as a list item. */
    private SiteSecurable librarySecurable(String driveItemId) {
        String number = getItem(driveItemId).getSharepointIds().getListItemId();
        check(!number.isEmpty(), "the item's list item id is unknown");
        return SiteSecurable.newBuilder().setListItem(ListItemRef.newBuilder()
                .setSiteId(siteId).setListId(folderSecurable().getListItem().getListId())
                .setItemId(number)).build();
    }

    private String roleAssignmentsOfTheSharedFile() {
        need(organizationLinkId, "the organization link");
        String unique = getItem(shared.getId()).getSharepointIds().getListItemUniqueId();
        check(!unique.isEmpty(), "the shared file's unique id is unknown");
        SiteSecurable file = librarySecurable(shared.getId());
        sharedSecurable = file;
        List<SiteRoleAssignment> assignments = new ArrayList<>();
        List<SiteRoleAssignment> onTheSite = new ArrayList<>();
        List<SiteGroup> groups = new ArrayList<>();
        try {
            check(stub().getSiteRoleInheritance(GetSiteRoleInheritanceRequest.newBuilder()
                    .setSecurable(file).build()).getHasUniqueRoleAssignments(),
                    "a shared file still inherits its role assignments");
            stub().listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder().setSecurable(file)
                    .build()).forEachRemaining(response -> assignments.add(response.getAssignment()));
            stub().listSiteRoleAssignments(ListSiteRoleAssignmentsRequest.newBuilder()
                    .setSecurable(SiteSecurable.newBuilder().setSiteId(siteId)).build())
                    .forEachRemaining(response -> onTheSite.add(response.getAssignment()));
            stub().listSiteGroups(ListSiteGroupsRequest.newBuilder().setSiteId(siteId).build())
                    .forEachRemaining(response -> groups.add(response.getGroup()));
        } catch (StatusRuntimeException e) {
            throw notGranted(e);
        }
        Map<String, Set<String>> levels = new TreeMap<>();
        for (SiteRoleAssignment assignment : assignments) {
            checkPrincipal(assignment.getMember());
            if (assignment.getMember().hasSharingLink()) {
                SiteSharingLinkGroup link = assignment.getMember().getSharingLink();
                check(link.getItemUniqueId().equalsIgnoreCase(unique),
                        "the file's role assignments name the link group of another item");
                check(!link.getKind().isEmpty() && !link.getShareId().isEmpty(),
                        "a link group has no kind or no share id");
                assignment.getRolesList().forEach(role -> levels
                        .computeIfAbsent(link.getScope() + " (" + link.getKind() + ")", key -> new TreeSet<>())
                        .add(role.getName()));
            }
        }
        check(levels.keySet().stream().anyMatch(scope -> scope.startsWith("organization ")),
                "no role assignment names the group of the organization link: " + levels.keySet());
        check(usersLinkId == null || levels.keySet().stream().anyMatch(scope -> scope.startsWith("users ")),
                "no role assignment names the group of the link for named people: " + levels.keySet());
        Set<String> siteLevels = new TreeSet<>();
        Set<String> groupOnTheSite = new TreeSet<>();
        int limited = 0;
        for (SiteRoleAssignment assignment : onTheSite) {
            if (assignment.getMember().getSharingLink().getItemUniqueId().equalsIgnoreCase(unique)) {
                assignment.getRolesList().forEach(role -> siteLevels.add(role.getName()));
            }
            if (groupId != null && assignment.getMember().getDirectoryObjectId().equalsIgnoreCase(groupId)) {
                assignment.getRolesList().forEach(role -> groupOnTheSite.add(role.getName()));
            }
            // Role type 1 is SharePoint's Guest type: Limited Access.
            limited += assignment.getRolesCount() > 0 && assignment.getRolesList().stream()
                    .allMatch(role -> role.getRoleTypeKind() == 1) ? 1 : 0;
        }
        Set<String> groupOnTheFile = new TreeSet<>();
        for (SiteRoleAssignment assignment : assignments) {
            if (groupId != null && assignment.getMember().getDirectoryObjectId().equalsIgnoreCase(groupId)) {
                assignment.getRolesList().forEach(role -> groupOnTheFile.add(role.getName()));
            }
        }
        check(invitedPermissionId == null || !groupOnTheFile.isEmpty(),
                "the file's role assignments do not name the invited group");
        int linkGroups = 0;
        int members = 0;
        boolean userIsMember = false;
        for (SiteGroup group : groups) {
            if (!group.getSharingLink().getItemUniqueId().equalsIgnoreCase(unique)) {
                continue;
            }
            linkGroups++;
            if (!group.getSharingLink().getScope().equals("users")) {
                continue;
            }
            List<SitePrincipal> people = new ArrayList<>();
            stub().listSiteGroupMembers(ListSiteGroupMembersRequest.newBuilder().setSiteId(siteId)
                    .setGroupId(group.getId()).build())
                    .forEachRemaining(response -> people.add(response.getMember()));
            people.forEach(WriteScenarios::checkPrincipal);
            members += people.size();
            userIsMember |= userPrincipalName != null && people.stream().anyMatch(person ->
                    person.getUserPrincipalName().equalsIgnoreCase(userPrincipalName));
        }
        check(linkGroups >= levels.size(), "the site lists " + linkGroups + " link group(s) of the file for "
                + levels.size() + " link(s) in its role assignments");
        check(usersLinkId == null || userPrincipalName == null || userIsMember,
                "the group of the link for named people does not hold the user it was granted to");
        return "the file has its own role assignments; link groups and their levels " + levels + "; "
                + linkGroups + " link group(s) of the file among the site's groups, " + members
                + " member(s) in the one for named people; on the site those groups hold "
                + (siteLevels.isEmpty() ? "nothing" : siteLevels)
                + (invitedPermissionId == null ? "" : "; the invited group holds " + groupOnTheFile
                        + " on the file and " + (groupOnTheSite.isEmpty() ? "nothing" : groupOnTheSite)
                        + " on the site")
                + "; " + limited + " assignment(s) on the site hold only Limited Access";
    }

    private String revokeLinkAccess() throws InterruptedException {
        need(usersLinkId, "the link for named people");
        Permission after = stub().revokeLinkAccess(RevokeLinkAccessRequest.newBuilder()
                .setItem(ref(shared.getId())).setPermissionId(usersLinkId).addGrantees(theUser()).build())
                .getPermission();
        check(after.getId().equals(usersLinkId), "the answer is another permission");
        int reads = 1;
        while (names(permissionOfTheSharedFile(usersLinkId), userId)) {
            check(reads < DIRECTORY_READS, "after " + reads + " reads the link still names the user");
            reads++;
            pause();
        }
        return "the answer names " + after.getGrantedToIdentitiesCount()
                + " grantee(s); the link read back no longer names the user on read " + reads;
    }

    private String deletePermissions() {
        need(organizationLinkId, "the organization link");
        List<String> ids = new ArrayList<>();
        for (String id : Arrays.asList(invitedPermissionId, usersLinkId, anyoneLinkId, organizationLinkId)) {
            if (id != null) {
                ids.add(id);
            }
        }
        for (String id : ids) {
            stub().deletePermission(DeletePermissionRequest.newBuilder().setTarget(sharedTarget())
                    .setPermissionId(id).build());
        }
        String invited = invitedPermissionId;
        invitedPermissionId = null;
        usersLinkId = null;
        anyoneLinkId = null;
        Map<String, Permission> left = new HashMap<>();
        int[] links = {0};
        stub().listPermissions(ListPermissionsRequest.newBuilder().setTarget(sharedTarget()).build())
                .forEachRemaining(response -> {
                    left.put(response.getPermission().getId(), response.getPermission());
                    links[0] += response.getPermission().hasLink() ? 1 : 0;
                });
        check(links[0] == 0, links[0] + " link(s) are still in the ACL");
        left.keySet().retainAll(ids);
        String user = "";
        if (invited != null && left.containsKey(invited)) {
            // The group held a permission of its own before the invitation; that one stays.
            Set<String> roles = new TreeSet<>(left.remove(invited).getRolesList());
            check(!rolesBeforeTheInvitation.isEmpty() && roles.equals(rolesBeforeTheInvitation),
                    "the invited group is still in the ACL with " + roles + "; before the invitation "
                            + rolesBeforeTheInvitation);
            user = "; the group is back to " + roles + ", as before the invitation";
        }
        check(left.isEmpty(), left.size() + " removed permission(s) are still in the ACL");
        return ids.size() + " permission(s) removed; the ACL lists no link" + user;
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

    /** Removes the test folder, the test list and a group the run left behind, and nothing else. */
    private String cleanUp() throws InterruptedException {
        List<String> lines = new ArrayList<>();
        boolean clean = true;
        if (entraGroupId != null) {
            try {
                onceKnown(() -> stub().deleteGroup(ai.pipestream.microsoft.v1.DeleteGroupRequest.newBuilder()
                        .setGroupId(entraGroupId).build()));
                lines.add("Entra group " + entraGroupName() + " deleted");
            } catch (StatusRuntimeException e) {
                clean = false;
                lines.add("Entra group " + entraGroupName() + " NOT CLEANED UP, " + e.getStatus());
            }
        }
        if (ownGroupId != null) {
            try {
                stub().deleteSiteGroup(ai.pipestream.microsoft.v1.DeleteSiteGroupRequest.newBuilder().setSiteId(siteId)
                        .setGroupId(ownGroupId).build());
                lines.add("group " + listName + "-group deleted");
            } catch (StatusRuntimeException e) {
                clean = false;
                lines.add("group " + listName + "-group NOT CLEANED UP, " + e.getStatus());
            }
        }
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
