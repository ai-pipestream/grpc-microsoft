package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ConflictBehavior;
import ai.pipestream.microsoft.v1.CreateFolderRequest;
import ai.pipestream.microsoft.v1.CreateFolderResponse;
import ai.pipestream.microsoft.v1.CreateListItemRequest;
import ai.pipestream.microsoft.v1.CreateListItemResponse;
import ai.pipestream.microsoft.v1.DeleteDriveItemRequest;
import ai.pipestream.microsoft.v1.DeleteDriveItemResponse;
import ai.pipestream.microsoft.v1.DeleteListItemRequest;
import ai.pipestream.microsoft.v1.DeleteListItemResponse;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.FieldWriteState;
import ai.pipestream.microsoft.v1.FileSystemInfo;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemTarget;
import ai.pipestream.microsoft.v1.ListRef;
import ai.pipestream.microsoft.v1.RestoreDriveItemVersionRequest;
import ai.pipestream.microsoft.v1.RestoreDriveItemVersionResponse;
import ai.pipestream.microsoft.v1.UpdateDriveItemRequest;
import ai.pipestream.microsoft.v1.UpdateDriveItemResponse;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsRequest;
import ai.pipestream.microsoft.v1.UpdateListItemFieldsResponse;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;
import io.grpc.Status;

import java.io.IOException;
import java.time.Instant;
import java.util.List;
import java.util.Map;
import java.util.function.Supplier;

/** Builds the Graph request and the typed response of each batchable write. */
final class WritePlans {

    /**
     * Turns a Graph response into the rpc's response.
     *
     * @param <R> the response type
     */
    @FunctionalInterface
    interface ReplyMapper<R> {

        /**
         * Maps a response.
         *
         * @param reply the Graph response; {@code null} when the plan has no call
         * @return the rpc response
         */
        R map(GraphClient.Reply reply);
    }

    /**
     * Work that follows the Graph request of a plan.
     *
     * @param <R> the rpc response type
     */
    @FunctionalInterface
    interface After<R> {

        /**
         * Finishes the write.
         *
         * @param result the response built from the Graph answer
         * @return the final response
         * @throws IOException if a call returns an error
         * @throws InterruptedException if a call is interrupted
         */
        R apply(R result) throws IOException, InterruptedException;
    }

    /**
     * One write, ready to send.
     *
     * @param call the Graph request; {@code null} when nothing is sent
     * @param mapper builds the response
     * @param after what follows the Graph request; {@code null} when nothing does
     * @param <R> the response type
     */
    record Plan<R>(GraphClient.Call call, ReplyMapper<R> mapper, After<R> after) {

        Plan(GraphClient.Call call, ReplyMapper<R> mapper) {
            this(call, mapper, null);
        }

        /**
         * Builds the response from the Graph answer and does what follows.
         *
         * @param reply the Graph response; {@code null} when the plan has no call
         * @return the rpc response
         * @throws IOException if a following call returns an error
         * @throws InterruptedException if a following call is interrupted
         */
        R finish(GraphClient.Reply reply) throws IOException, InterruptedException {
            R result = mapper.map(reply);
            return after == null ? result : after.apply(result);
        }

        /**
         * Sends the request on its own.
         *
         * @param graph the client to send with
         * @return the rpc response
         * @throws IOException if Graph returns an error
         * @throws InterruptedException if the HTTP call is interrupted
         */
        R run(GraphClient graph) throws IOException, InterruptedException {
            return finish(call == null ? null : graph.execute(call));
        }
    }

    private final MicrosoftMapper mapper;
    private final Supplier<SharePointRest> sharePoint;

    WritePlans(MicrosoftMapper mapper, Supplier<SharePointRest> sharePoint) {
        this.mapper = mapper;
        this.sharePoint = sharePoint;
    }

    /**
     * Writes hyperlink values with SharePoint's own API.
     *
     * @param session the connection
     * @param listItemRoute Graph route of the list item
     * @param links the values
     * @throws IOException if a call returns an error
     * @throws InterruptedException if a call is interrupted
     */
    void writeLinks(MicrosoftSessions.Session session, String listItemRoute,
            List<ListFieldWriter.Link> links) throws IOException, InterruptedException {
        JsonNode ids = session.graph().get(listItemRoute + "?$select=id,sharepointIds", Map.of())
                .path("sharepointIds");
        String webUrl = ids.path("siteUrl").asText("");
        if (webUrl.isBlank()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("Graph names no site URL to write the hyperlink through")
                    .asRuntimeException();
        }
        ObjectNode body = GraphClient.object();
        ArrayNode values = body.putArray("formValues");
        for (ListFieldWriter.Link link : links) {
            // SharePoint's text form of a link: the URL with commas doubled, then the text.
            values.addObject().put("FieldName", link.column())
                    .put("FieldValue", link.url().replace(",", ",,") + ", " + link.description());
        }
        body.put("bNewDocumentUpdate", false);
        JsonNode answer = sharePoint.get().webAt(session, webUrl).post(SharePointRest.item(
                ids.path("listId").asText(""), ids.path("listItemId").asText(""))
                + "/ValidateUpdateListItem", body);
        for (JsonNode value : answer.path("value")) {
            if (value.path("HasException").asBoolean(false)) {
                throw Status.INVALID_ARGUMENT.withDescription("column "
                        + value.path("FieldName").asText("") + ": "
                        + value.path("ErrorMessage").asText("")).asRuntimeException();
            }
        }
    }

    /**
     * The Graph {@code conflictBehavior} value.
     *
     * @param behavior the requested behaviour
     * @return {@code fail}, {@code replace} or {@code rename}; empty when unspecified
     */
    static String conflict(ConflictBehavior behavior) {
        return switch (behavior) {
            case CONFLICT_BEHAVIOR_FAIL -> "fail";
            case CONFLICT_BEHAVIOR_REPLACE -> "replace";
            case CONFLICT_BEHAVIOR_RENAME -> "rename";
            case CONFLICT_BEHAVIOR_UNSPECIFIED -> "";
            case UNRECOGNIZED -> throw new IllegalArgumentException("unknown conflict_behavior");
        };
    }

    Plan<CreateFolderResponse> createFolder(CreateFolderRequest request) {
        String driveId = request.getDriveId();
        GraphPaths.childByName(driveId, "", "", request.getName());
        ObjectNode body = GraphClient.object();
        body.put("name", request.getName());
        body.putObject("folder");
        String conflict = conflict(request.getConflictBehavior());
        if (!conflict.isEmpty()) {
            body.put("@microsoft.graph.conflictBehavior", conflict);
        }
        String route = GraphPaths.children(driveId, request.getParentId(), request.getParentPath());
        return new Plan<>(new GraphClient.Call("POST", route, Map.of(), body),
                reply -> CreateFolderResponse.newBuilder()
                        .setItem(valid(mapper.toDriveItem(reply.body(), driveId))).build());
    }

    Plan<UpdateDriveItemResponse> updateDriveItem(MicrosoftSessions.Session session,
            UpdateDriveItemRequest request) throws IOException, InterruptedException {
        DriveItemRef ref = request.getItem();
        if (request.hasDescription() && !request.getDescription().isEmpty()) {
            requirePersonalDrive(session, ref.getDriveId());
        }
        ObjectNode body = GraphClient.object();
        if (request.hasName()) {
            GraphPaths.childByName(ref.getDriveId(), "", "", request.getName());
            body.put("name", request.getName());
        }
        if (!request.getNewParentId().isBlank()) {
            body.putObject("parentReference").put("id", request.getNewParentId());
        }
        if (request.hasDescription()) {
            body.put("description", request.getDescription());
        }
        if (request.hasFileSystemInfo()) {
            FileSystemInfo info = request.getFileSystemInfo();
            ObjectNode times = body.putObject("fileSystemInfo");
            if (info.hasCreatedAt()) {
                times.put("createdDateTime", instant(info.getCreatedAt()));
            }
            if (info.hasLastModifiedAt()) {
                times.put("lastModifiedDateTime", instant(info.getLastModifiedAt()));
            }
        }
        if (body.isEmpty()) {
            throw new IllegalArgumentException("the request changes nothing");
        }
        String route = GraphPaths.item(ref.getDriveId(), ref.getItemId());
        String conflict = conflict(request.getConflictBehavior());
        if (!conflict.isEmpty()) {
            route += "?@microsoft.graph.conflictBehavior=" + conflict;
        }
        return new Plan<>(
                new GraphClient.Call("PATCH", route, ifMatch(request.getExpectedEtag()), body),
                reply -> UpdateDriveItemResponse.newBuilder()
                        .setItem(valid(mapper.toDriveItem(reply.body(), ref.getDriveId()))).build());
    }

    /** Fails with FAILED_PRECONDITION unless the drive is a OneDrive personal drive. */
    private static void requirePersonalDrive(MicrosoftSessions.Session session, String driveId)
            throws IOException, InterruptedException {
        JsonNode drive = session.graph().get(
                GraphPaths.drive(driveId) + "?$select=id,driveType", Map.of());
        String type = drive.path("driveType").asText("");
        if (!"personal".equals(type)) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("description is stored on OneDrive personal drives only; drive "
                            + driveId + " is " + (type.isEmpty() ? "of unknown type" : type))
                    .asRuntimeException();
        }
    }

    Plan<DeleteDriveItemResponse> deleteDriveItem(DeleteDriveItemRequest request) {
        DriveItemRef ref = request.getItem();
        String route = GraphPaths.item(ref.getDriveId(), ref.getItemId());
        GraphClient.Call call = request.getPermanent()
                ? new GraphClient.Call("POST", route + "/permanentDelete",
                        ifMatch(request.getExpectedEtag()), GraphClient.object())
                : new GraphClient.Call("DELETE", route, ifMatch(request.getExpectedEtag()), null);
        return new Plan<>(call, reply -> DeleteDriveItemResponse.newBuilder().setItem(ref).build());
    }

    Plan<RestoreDriveItemVersionResponse> restoreDriveItemVersion(
            RestoreDriveItemVersionRequest request) {
        DriveItemRef ref = request.getItem();
        String route = GraphPaths.item(ref.getDriveId(), ref.getItemId()) + "/versions/"
                + GraphPaths.segment(request.getVersionId(), "version_id") + "/restoreVersion";
        return new Plan<>(
                new GraphClient.Call("POST", route, ifMatch(request.getExpectedEtag()),
                        GraphClient.object()),
                reply -> RestoreDriveItemVersionResponse.newBuilder().setItem(ref)
                        .setVersionId(request.getVersionId()).build());
    }

    Plan<CreateListItemResponse> createListItem(MicrosoftSessions.Session session,
            CreateListItemRequest request) throws IOException, InterruptedException {
        ListRef list = request.getList();
        String route = GraphPaths.list(list.getSiteId(), list.getListId());
        ListFieldWriter.Patch patch = ListFieldWriter.translate(request.getFieldsList(),
                lookups(session, route, list.getSiteId()), FieldWriteState.FIELD_WRITE_STATE_WRITTEN);
        ObjectNode body = GraphClient.object();
        if (!request.getContentTypeId().isBlank()) {
            body.putObject("contentType").put("id", request.getContentTypeId());
        }
        body.set("fields", patch.fields());
        After<CreateListItemResponse> after = patch.links().isEmpty() ? null : created -> {
            String itemRoute = route + "/items/"
                    + GraphPaths.segment(created.getItem().getId(), "item_id");
            writeLinks(session, itemRoute, patch.links());
            return CreateListItemResponse.newBuilder()
                    .setItem(valid(mapper.toListItem(
                            session.graph().get(itemRoute + "?$expand=fields", Map.of()),
                            list.getSiteId(), list.getListId())))
                    .build();
        };
        return new Plan<>(new GraphClient.Call("POST", route + "/items", Map.of(), body),
                reply -> CreateListItemResponse.newBuilder()
                        .setItem(valid(mapper.toListItem(reply.body(), list.getSiteId(),
                                list.getListId())))
                        .build(), after);
    }

    Plan<UpdateListItemFieldsResponse> updateListItemFields(MicrosoftSessions.Session session,
            UpdateListItemFieldsRequest request) throws IOException, InterruptedException {
        ListItemTarget target = request.getPatch().getTarget();
        String fieldsRoute;
        ListFieldWriter.Lookups lookups;
        switch (target.getTargetCase()) {
            case LIST_ITEM -> {
                ListItemRef ref = target.getListItem();
                fieldsRoute = GraphPaths.listItem(ref.getSiteId(), ref.getListId(), ref.getItemId())
                        + "/fields";
                lookups = lookups(session, GraphPaths.list(ref.getSiteId(), ref.getListId()),
                        ref.getSiteId());
            }
            case DRIVE_ITEM -> {
                DriveItemRef ref = target.getDriveItem();
                fieldsRoute = GraphPaths.item(ref.getDriveId(), ref.getItemId())
                        + "/listItem/fields";
                lookups = driveLookups(session, ref.getDriveId());
            }
            default -> throw new IllegalArgumentException("the patch needs a target");
        }
        if (request.getPatch().getFieldsCount() == 0) {
            throw new IllegalArgumentException("the patch names no columns");
        }
        boolean dryRun = request.getDryRun();
        ListFieldWriter.Patch patch = ListFieldWriter.translate(request.getPatch().getFieldsList(),
                lookups, dryRun ? FieldWriteState.FIELD_WRITE_STATE_SKIPPED
                        : FieldWriteState.FIELD_WRITE_STATE_WRITTEN);
        if (dryRun) {
            return new Plan<>(null, reply -> UpdateListItemFieldsResponse.newBuilder()
                    .addAllOutcomes(patch.outcomes()).build());
        }
        String itemRoute = fieldsRoute.substring(0, fieldsRoute.length() - "/fields".length());
        After<UpdateListItemFieldsResponse> after = patch.links().isEmpty() ? null : written -> {
            writeLinks(session, itemRoute, patch.links());
            JsonNode fields = session.graph().get(fieldsRoute, Map.of());
            return UpdateListItemFieldsResponse.newBuilder()
                    .addAllOutcomes(patch.outcomes())
                    .addAllFields(mapper.toListColumns(fields))
                    .setEtag(fields.path("@odata.etag").asText(""))
                    .build();
        };
        if (patch.fields().isEmpty()) {
            // Only hyperlinks: nothing for Graph to write, so the eTag is checked by reading.
            String expected = request.getExpectedEtag();
            return new Plan<>(
                    new GraphClient.Call("GET", itemRoute + "?$select=id,eTag", Map.of(), null),
                    reply -> {
                        if (!expected.isEmpty()
                                && !expected.equals(reply.body().path("eTag").asText(""))
                                && !expected.equals(reply.body().path("@odata.etag").asText(""))) {
                            throw Status.ABORTED.withDescription("the item changed: its eTag is "
                                    + "not " + expected).asRuntimeException();
                        }
                        return UpdateListItemFieldsResponse.newBuilder()
                                .addAllOutcomes(patch.outcomes()).build();
                    }, after);
        }
        return new Plan<>(
                new GraphClient.Call("PATCH", fieldsRoute, ifMatch(request.getExpectedEtag()),
                        patch.fields()),
                reply -> UpdateListItemFieldsResponse.newBuilder()
                        .addAllOutcomes(patch.outcomes())
                        .addAllFields(mapper.toListColumns(reply.body()))
                        .setEtag(reply.body().path("@odata.etag").asText(""))
                        .build(), after);
    }

    Plan<DeleteListItemResponse> deleteListItem(DeleteListItemRequest request) {
        ListItemRef ref = request.getItem();
        String route = GraphPaths.listItem(ref.getSiteId(), ref.getListId(), ref.getItemId());
        return new Plan<>(
                new GraphClient.Call("DELETE", route, ifMatch(request.getExpectedEtag()), null),
                reply -> DeleteListItemResponse.newBuilder().setItem(ref).build());
    }

    /**
     * Lookups for the list behind a drive.
     *
     * @param session the connection
     * @param driveId the drive id
     * @return the lookups
     */
    ListFieldWriter.Lookups driveLookups(MicrosoftSessions.Session session, String driveId) {
        String listRoute = GraphPaths.drive(driveId) + "/list";
        return new ListFieldWriter.Lookups() {
            @Override
            public int personId(String email) throws IOException, InterruptedException {
                JsonNode list = session.graph().get(
                        listRoute + "?$select=id,parentReference,sharepointIds", Map.of());
                String siteId = list.path("parentReference").path("siteId").asText("");
                if (siteId.isBlank()) {
                    siteId = list.path("sharepointIds").path("siteId").asText("");
                }
                if (siteId.isBlank()) {
                    throw Status.FAILED_PRECONDITION
                            .withDescription("drive " + driveId + " names no site to look "
                                    + email + " up in")
                            .asRuntimeException();
                }
                return WritePlans.personId(session, siteId, email);
            }

            @Override
            public String termNoteColumn(String termColumn)
                    throws IOException, InterruptedException {
                return WritePlans.this.termNoteColumn(session, listRoute, termColumn);
            }
        };
    }

    private ListFieldWriter.Lookups lookups(MicrosoftSessions.Session session,
            String listRoute, String siteId) {
        return new ListFieldWriter.Lookups() {
            @Override
            public int personId(String email) throws IOException, InterruptedException {
                return WritePlans.personId(session, siteId, email);
            }

            @Override
            public String termNoteColumn(String termColumn)
                    throws IOException, InterruptedException {
                return WritePlans.this.termNoteColumn(session, listRoute, termColumn);
            }
        };
    }

    private static int personId(MicrosoftSessions.Session session, String siteId, String email)
            throws IOException, InterruptedException {
        if (email.isBlank()) {
            throw new IllegalArgumentException("a person needs lookup_id or email");
        }
        String route = "/sites/" + GraphPaths.site(siteId) + "/lists/"
                + GraphPaths.segment("User Information List", "list_id")
                + "/items?$select=id&$top=1&$filter="
                + GraphPaths.query("fields/EMail eq '" + email.replace("'", "''") + "'");
        JsonNode page = session.graph().get(route,
                Map.of("Prefer", GraphPaths.HONOR_NON_INDEXED));
        String id = page.path("value").path(0).path("id").asText("");
        try {
            return Integer.parseInt(id);
        } catch (NumberFormatException e) {
            throw Status.NOT_FOUND
                    .withDescription("the site has no user with email " + email)
                    .asRuntimeException();
        }
    }

    private String termNoteColumn(MicrosoftSessions.Session session, String listRoute,
            String termColumn) throws IOException, InterruptedException {
        Map<String, String> displayNames = new java.util.HashMap<>();
        Map<String, String> namesByDisplayName = new java.util.HashMap<>();
        session.pager().forEach(
                java.util.List.of(listRoute + "/columns?$select=name,displayName"), Map.of(),
                node -> {
                    String name = node.path("name").asText("");
                    String displayName = node.path("displayName").asText("");
                    displayNames.put(name, displayName);
                    namesByDisplayName.putIfAbsent(displayName, name);
                    return true;
                });
        String displayName = displayNames.get(termColumn);
        String note = displayName == null ? null : namesByDisplayName.get(displayName + "_0");
        if (note == null) {
            note = termNoteColumnFromSharePoint(session, listRoute, termColumn);
        }
        if (note.isBlank()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("no hidden note column found for term column " + termColumn
                            + "; set term_note_column")
                    .asRuntimeException();
        }
        return note;
    }

    /** The note column SharePoint names for a term column; empty when it names none. */
    private String termNoteColumnFromSharePoint(MicrosoftSessions.Session session,
            String listRoute, String termColumn) throws IOException, InterruptedException {
        JsonNode ids = session.graph().get(listRoute + "?$select=id,sharepointIds", Map.of())
                .path("sharepointIds");
        String webUrl = ids.path("siteUrl").asText("");
        if (webUrl.isBlank()) {
            return "";
        }
        SharePointRest.Web web = sharePoint.get().webAt(session, webUrl);
        String fields = SharePointRest.fields(ids.path("listId").asText(""));
        String noteId = web.get(fields + "/getbyinternalnameortitle('"
                + GraphPaths.query(termColumn.replace("'", "''")) + "')").path("TextField").asText("");
        if (noteId.isBlank()) {
            return "";
        }
        return web.get(fields + "(guid'" + GraphPaths.segment(noteId, "note column id")
                + "')?$select=InternalName").path("InternalName").asText("");
    }

    private static <T extends com.google.protobuf.Message> T valid(T message) {
        MicrosoftValidator.create().requireValid(message);
        return message;
    }

    private static Map<String, String> ifMatch(String etag) {
        return etag == null || etag.isBlank() ? Map.of() : Map.of("If-Match", etag);
    }

    private static String instant(Timestamp timestamp) {
        return Instant.ofEpochSecond(timestamp.getSeconds(), timestamp.getNanos()).toString();
    }
}
