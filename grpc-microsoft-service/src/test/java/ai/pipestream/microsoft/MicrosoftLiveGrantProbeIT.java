package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.CreateSharingLinkRequest;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeleteListItemRequest;
import ai.pipestream.microsoft.v1.DeletePermissionRequest;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.EnsureListRequest;
import ai.pipestream.microsoft.v1.EnsureListResponse;
import ai.pipestream.microsoft.v1.EnsureState;
import ai.pipestream.microsoft.v1.FieldValue;
import ai.pipestream.microsoft.v1.GetPermissionRequest;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.ListItemPatch;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemTarget;
import ai.pipestream.microsoft.v1.ListListItemsRequest;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.ListTemplate;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MutateRequest;
import ai.pipestream.microsoft.v1.MutateResponse;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.UpdateDriveItemRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsRequest;
import ai.pipestream.microsoft.v1.UploadItemEnd;
import ai.pipestream.microsoft.v1.UploadItemHeader;
import ai.pipestream.microsoft.v1.UploadItemRequest;
import ai.pipestream.microsoft.v1.UploadItemResponse;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import com.google.protobuf.ByteString;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Tries the writes that need no new list against the list and library the
 * fixture left on the test site, and prints whether each was allowed. A
 * refused call is an answer, not a failure, so the same run shows what a
 * lower site grant permits. It removes what it wrote, and what an earlier
 * run could not remove.
 *
 * <p>Runs only when {@code MICROSOFT_TENANT_ID}, {@code MICROSOFT_CLIENT_ID},
 * {@code MICROSOFT_CLIENT_SECRET} and {@code MICROSOFT_SITE_ID} are set and
 * {@code MICROSOFT_LIVE_GRANT_PROBE=1}. Run with
 * {@code ./gradlew :grpc-microsoft-service:liveGrantProbeTest}.</p>
 */
class MicrosoftLiveGrantProbeIT {

    /** Environment switch that allows the probe to write to the tenant. */
    static final String ENV_LIVE_GRANT_PROBE = "MICROSOFT_LIVE_GRANT_PROBE";

    /** Prefix of every file the probe uploads. */
    static final String FILE_PREFIX = "pipestream-grant-probe-";

    /** Prefix of the title of every list item the probe creates. */
    static final String TITLE_PREFIX = "grant probe";

    private static final String TAG = "[MicrosoftLiveGrantProbeIT] ";

    private final List<String> unexpected = new ArrayList<>();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private MicrosoftServiceGrpc.MicrosoftServiceStub async;
    private String siteId;
    private String driveId;
    private String listId;

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void writesNextToTheFixtureAreAllowedOrRefused() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_SITE_ID)),
                "no live Microsoft application credentials and site in the environment; skipping");
        assumeTrue("1".equals(System.getenv(ENV_LIVE_GRANT_PROBE)),
                ENV_LIVE_GRANT_PROBE + " is not 1; not writing to the tenant");
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.fromEnvironment();
        GraphAuth auth = new GraphAuth(config.authConfig());
        String token = auth.clientCredentials().accessToken();
        String graphBaseUrl = GraphLinks.stripTrailingSlashes(config.graphBaseUrl());
        GraphClient client = new GraphClient(graphBaseUrl, () -> token, AuthMode.AUTH_MODE_APPLICATION);
        String site = "/sites/" + GraphPaths.site(config.siteId());
        siteId = client.get(site + "?$select=id").path("id").asText("");
        driveId = config.driveIds().isEmpty()
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
        List<String> left;
        try {
            stub = MicrosoftServiceGrpc.newBlockingStub(channel);
            async = MicrosoftServiceGrpc.newStub(channel);
            EnsureListResponse list = stub.ensureList(EnsureListRequest.newBuilder().setSiteId(siteId)
                    .setDisplayName(MicrosoftLiveFixtureIT.LIST_NAME)
                    .setTemplate(ListTemplate.LIST_TEMPLATE_GENERIC_LIST).setDryRun(true).build());
            assertThat(list.getState()).as("the fixture list; run liveFixtureTest first")
                    .isEqualTo(EnsureState.ENSURE_STATE_EXISTS);
            listId = list.getList().getId();
            System.out.println(TAG + "left by an earlier run: " + sweep());
            listItemWrites();
            fileWrites();
        } finally {
            try {
                left = stub == null ? List.of() : sweep();
            } finally {
                channel.shutdownNow();
                server.shutdownNow();
                server.awaitTermination(10, TimeUnit.SECONDS);
                executor.shutdownNow();
            }
        }
        System.out.println(TAG + "cleanup: " + left);
        assertThat(left).as("what the probe could not remove").noneMatch(line -> line.contains("NOT"));
        assertThat(unexpected).as("answers other than OK and PERMISSION_DENIED").isEmpty();
    }

    private void listItemWrites() throws Exception {
        ListRef list = ListRef.newBuilder().setSiteId(siteId).setListId(listId).build();
        ListItem[] created = new ListItem[1];
        row("CreateListItem", () -> created[0] = stub.createListItem(CreateListItemRequest.newBuilder()
                .setList(list).addFields(title(TITLE_PREFIX)).build()).getItem());
        // A refused create leaves the fixture's own item to try the other writes on, with its own title.
        String itemId = created[0] != null ? created[0].getId() : stub.listListItems(
                ListListItemsRequest.newBuilder().setSiteId(siteId).setListId(listId).build())
                .next().getItem().getId();
        String retitled = created[0] != null ? TITLE_PREFIX + " retitled" : "coverage sample";
        ListItemTarget target = ListItemTarget.newBuilder().setListItem(ref(itemId)).build();
        row("UpdateListItemFields", () -> stub.updateListItemFields(UpdateListItemFieldsRequest.newBuilder()
                .setPatch(ListItemPatch.newBuilder().setTarget(target).addFields(title(retitled))).build()));
        mutate(List.of(
                MutateRequest.newBuilder().setOpId("create list item").setCreateListItem(
                        CreateListItemRequest.newBuilder().setList(list)
                                .addFields(title(TITLE_PREFIX + " batch"))).build(),
                MutateRequest.newBuilder().setOpId("update list item fields").setUpdateListItemFields(
                        UpdateListItemFieldsRequest.newBuilder().setPatch(ListItemPatch.newBuilder()
                                .setTarget(target).addFields(title(retitled)))).build()));
        if (created[0] != null) {
            row("DeleteListItem", () -> stub.deleteListItem(
                    DeleteListItemRequest.newBuilder().setItem(ref(itemId)).build()));
        } else {
            System.out.println(TAG + "| DeleteListItem | NOT TRIED, no item of the probe's own");
        }
    }

    private void fileWrites() throws Exception {
        String name = Long.toString(System.currentTimeMillis(), 36) + ".txt";
        DriveItem[] uploaded = new DriveItem[1];
        row("UploadItem", () -> uploaded[0] = upload(FILE_PREFIX + name));
        String fileId = uploaded[0] != null ? uploaded[0].getId() : fixtureFileId();
        PermissionTarget file = PermissionTarget.newBuilder()
                .setDriveItem(DriveItemRef.newBuilder().setDriveId(driveId).setItemId(fileId)).build();
        String[] link = new String[1];
        row("CreateSharingLink [view, organization]", () -> link[0] = stub.createSharingLink(
                CreateSharingLinkRequest.newBuilder().setTarget(file).setType("view")
                        .setScope("organization").build()).getPermission().getId());
        if (link[0] != null) {
            row("GetPermission", () -> stub.getPermission(GetPermissionRequest.newBuilder()
                    .setTarget(file).setPermissionId(link[0]).build()));
            row("DeletePermission", () -> stub.deletePermission(DeletePermissionRequest.newBuilder()
                    .setTarget(file).setPermissionId(link[0]).build()));
        } else {
            System.out.println(TAG + "| GetPermission, DeletePermission | NOT TRIED, no link");
        }
        if (uploaded[0] == null) {
            System.out.println(TAG + "| UpdateDriveItem, DeleteDriveItem | NOT TRIED, no file of the probe's own");
            return;
        }
        DriveItemRef ref = DriveItemRef.newBuilder().setDriveId(driveId).setItemId(fileId).build();
        row("UpdateDriveItem [name]", () -> stub.updateDriveItem(UpdateDriveItemRequest.newBuilder()
                .setItem(ref).setName(FILE_PREFIX + "renamed-" + name).build()));
        mutate(List.of(MutateRequest.newBuilder().setOpId("update drive item").setUpdateDriveItem(
                UpdateDriveItemRequest.newBuilder().setItem(ref).setName(FILE_PREFIX + "batch-" + name)).build()));
        row("DeleteDriveItem", () -> stub.deleteDriveItem(
                DeleteDriveItemRequest.newBuilder().setItem(ref).build()));
    }

    /** Prints one row for a call; anything but OK and PERMISSION_DENIED fails the run at its end. */
    private void row(String label, Callable<?> call) throws Exception {
        String outcome = "OK";
        try {
            call.call();
        } catch (io.grpc.StatusRuntimeException e) {
            outcome = e.getStatus().getCode().name();
            if (e.getStatus().getCode() != Status.Code.PERMISSION_DENIED) {
                unexpected.add(label + " " + outcome + " " + e.getStatus().getDescription());
            }
        }
        System.out.println(TAG + "| " + label + " | " + outcome);
    }

    /** Sends the ops in one Mutate call and prints one row for each. */
    private void mutate(List<MutateRequest> ops) throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<Throwable> error = new AtomicReference<>();
        List<MutateResponse> answers = new ArrayList<>();
        StreamObserver<MutateRequest> requests = async.mutate(new StreamObserver<>() {
            @Override
            public void onNext(MutateResponse value) {
                synchronized (answers) {
                    answers.add(value);
                }
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
        ops.forEach(requests::onNext);
        requests.onCompleted();
        assertThat(done.await(2, TimeUnit.MINUTES)).as("Mutate ended in time").isTrue();
        assertThat(error.get()).as("the Mutate call").isNull();
        assertThat(answers).as("one answer for each op").hasSameSizeAs(ops);
        for (MutateResponse answer : answers) {
            String outcome = "OK";
            if (answer.hasError()) {
                WriteErrorCode code = answer.getError().getCode();
                outcome = code.name().substring("WRITE_ERROR_CODE_".length());
                if (code != WriteErrorCode.WRITE_ERROR_CODE_PERMISSION_DENIED) {
                    unexpected.add("Mutate " + answer.getOpId() + " " + outcome + " "
                            + answer.getError().getMessage());
                }
            }
            System.out.println(TAG + "| Mutate [" + answer.getOpId() + "] | " + outcome);
        }
    }

    private DriveItem upload(String name) throws Exception {
        byte[] content = "A file of the grant probe. Safe to delete.\n".getBytes(StandardCharsets.UTF_8);
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
                .setDriveId(driveId).setName(name).setMediaType("text/plain").setSize(content.length)
                .setConflictBehavior(ConflictBehavior.CONFLICT_BEHAVIOR_FAIL)).build());
        requests.onNext(UploadItemRequest.newBuilder().setChunk(ByteString.copyFrom(content)).build());
        requests.onNext(UploadItemRequest.newBuilder().setEnd(UploadItemEnd.newBuilder()
                .setSha256(HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content))))
                .build());
        requests.onCompleted();
        assertThat(done.await(2, TimeUnit.MINUTES)).as("the upload ended in time").isTrue();
        if (error.get() != null) {
            throw Status.fromThrowable(error.get()).asRuntimeException();
        }
        return last.get().getItem();
    }

    private String fixtureFileId() {
        String[] id = new String[1];
        stub.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId).build())
                .forEachRemaining(response -> {
                    if (MicrosoftLiveFixtureIT.FILE_NAME.equals(response.getItem().getName())) {
                        id[0] = response.getItem().getId();
                    }
                });
        assertThat(id[0]).as("the fixture file; run liveFixtureTest first").isNotNull();
        return id[0];
    }

    /** Removes the probe's files and list items; a line holding NOT names one that stays. */
    private List<String> sweep() {
        List<String> lines = new ArrayList<>();
        stub.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId).build())
                .forEachRemaining(response -> {
                    if (response.getItem().getName().startsWith(FILE_PREFIX)) {
                        lines.add(removed("file", () -> stub.deleteDriveItem(DeleteDriveItemRequest.newBuilder()
                                .setItem(DriveItemRef.newBuilder().setDriveId(driveId)
                                        .setItemId(response.getItem().getId())).build())));
                    }
                });
        stub.listListItems(ListListItemsRequest.newBuilder().setSiteId(siteId).setListId(listId).build())
                .forEachRemaining(response -> {
                    ListItem item = response.getItem();
                    boolean own = item.getFieldsList().stream().anyMatch(column -> column.getName().equals("Title")
                            && column.getValueCase() == ListColumn.ValueCase.STRING_VALUE
                            && column.getStringValue().startsWith(TITLE_PREFIX));
                    if (own) {
                        lines.add(removed("list item", () -> stub.deleteListItem(
                                DeleteListItemRequest.newBuilder().setItem(ref(item.getId())).build())));
                    }
                });
        return lines;
    }

    private static String removed(String what, Runnable delete) {
        try {
            delete.run();
            return what + " removed";
        } catch (io.grpc.StatusRuntimeException e) {
            return what + " NOT removed, " + e.getStatus().getCode();
        }
    }

    private ListItemRef ref(String itemId) {
        return ListItemRef.newBuilder().setSiteId(siteId).setListId(listId).setItemId(itemId).build();
    }

    private static FieldValue.Builder title(String text) {
        return FieldValue.newBuilder().setName("Title").setText(text);
    }
}
