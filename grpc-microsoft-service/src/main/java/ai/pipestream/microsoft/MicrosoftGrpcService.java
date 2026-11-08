package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ColumnDefinition;
import ai.pipestream.microsoft.v1.DeltaDriveItemsRequest;
import ai.pipestream.microsoft.v1.DeltaDriveItemsResponse;
import ai.pipestream.microsoft.v1.DeltaListItemsRequest;
import ai.pipestream.microsoft.v1.DeltaListItemsResponse;
import ai.pipestream.microsoft.v1.DeltaResync;
import ai.pipestream.microsoft.v1.DeltaSitesRequest;
import ai.pipestream.microsoft.v1.DeltaSitesResponse;
import ai.pipestream.microsoft.v1.DownloadContentRequest;
import ai.pipestream.microsoft.v1.DownloadContentResponse;
import ai.pipestream.microsoft.v1.DownloadItemRequest;
import ai.pipestream.microsoft.v1.DownloadItemResponse;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GetDriveRequest;
import ai.pipestream.microsoft.v1.GetDriveResponse;
import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.GetItemResponse;
import ai.pipestream.microsoft.v1.GetListItemRequest;
import ai.pipestream.microsoft.v1.GetListItemResponse;
import ai.pipestream.microsoft.v1.GetListRequest;
import ai.pipestream.microsoft.v1.GetListResponse;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.GetSiteRequest;
import ai.pipestream.microsoft.v1.GetSiteResponse;
import ai.pipestream.microsoft.v1.ItemReference;
import ai.pipestream.microsoft.v1.ListAllSitesRequest;
import ai.pipestream.microsoft.v1.ListAllSitesResponse;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListChildrenResponse;
import ai.pipestream.microsoft.v1.ListColumnsRequest;
import ai.pipestream.microsoft.v1.ListColumnsResponse;
import ai.pipestream.microsoft.v1.ListDrivesRequest;
import ai.pipestream.microsoft.v1.ListDrivesResponse;
import ai.pipestream.microsoft.v1.ListItemInclude;
import ai.pipestream.microsoft.v1.ListItemVersionsRequest;
import ai.pipestream.microsoft.v1.ListItemVersionsResponse;
import ai.pipestream.microsoft.v1.ListListItemsRequest;
import ai.pipestream.microsoft.v1.ListListItemsResponse;
import ai.pipestream.microsoft.v1.ListListsRequest;
import ai.pipestream.microsoft.v1.ListListsResponse;
import ai.pipestream.microsoft.v1.ListPermissionsRequest;
import ai.pipestream.microsoft.v1.ListPermissionsResponse;
import ai.pipestream.microsoft.v1.ListSitesRequest;
import ai.pipestream.microsoft.v1.ListSitesResponse;
import ai.pipestream.microsoft.v1.ListSubsitesRequest;
import ai.pipestream.microsoft.v1.ListSubsitesResponse;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.ProbeConnectionRequest;
import ai.pipestream.microsoft.v1.ProbeConnectionResponse;
import ai.pipestream.microsoft.v1.SharePointList;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.sync.v1.ConnectionServiceGrpc;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.ByteString;
import com.google.protobuf.Message;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.stub.StreamObserver;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * {@code MicrosoftService} over Microsoft Graph v1.0: every rpc of the contract,
 * one Graph route each (named on the rpc), with REST paging hidden behind
 * streams, delta links surfaced as typed terminal events, and binaries streamed
 * in chunks. Every outgoing entity is checked by {@link MicrosoftValidator}.
 * Requests that name a {@code connection_id} run against that catalog
 * connection ({@link MicrosoftSessions}); every client-supplied delta link or
 * resume cursor is validated against the configured Graph host and the route it
 * claims before the server's token is sent to it.
 */
public final class MicrosoftGrpcService
        implements MicrosoftServiceGrpc.AsyncService, io.grpc.BindableService {

    /**
     * The service definition, with every server-streaming rpc behind
     * {@link FlowControlInterceptor} (isReady-aware sends that stop on
     * cancellation).
     *
     * @return the bound service
     */
    @Override
    public io.grpc.ServerServiceDefinition bindService() {
        return io.grpc.ServerInterceptors.intercept(MicrosoftServiceGrpc.bindService(this),
                new FlowControlInterceptor());
    }


    /** Default cap for inlined file bytes: 25 MiB. */
    public static final long DEFAULT_ATTACHMENT_MAX_BYTES = 25L * 1024 * 1024;

    /**
     * Server default for {@code DownloadContentRequest.chunk_bytes = 0}: 1 MiB per
     * streamed message, well under gRPC's default 4 MiB message limit. Override per
     * process with {@link MicrosoftServer#ENV_DOWNLOAD_CHUNK_BYTES}.
     */
    public static final int DEFAULT_DOWNLOAD_CHUNK_BYTES = 1024 * 1024;

    /** Site ids returned by {@code ProbeConnection} when its limit is 0. */
    public static final int DEFAULT_PROBE_SITES = 5;

    private static final MicrosoftValidator VALIDATOR = MicrosoftValidator.create();

    private final MicrosoftSessions sessions;
    private final MicrosoftMapper mapper;
    private final long attachmentMaxBytes;
    private final MicrosoftChangeSink downstream;
    private final int downloadChunkBytes;

    /**
     * Creates a service with no downstream {@link MicrosoftChangeSink}.
     *
     * @param config connector config
     * @param files the authorized files API
     * @param attachmentMaxBytes server-wide inline byte limit used when a request
     *        leaves its own limit at 0; 0 = no limit
     */
    public MicrosoftGrpcService(MicrosoftConnectorConfig config, GraphFiles files,
            long attachmentMaxBytes) {
        this(config, files, attachmentMaxBytes, null);
    }

    /**
     * Creates a service. When {@code downstream} is non-null, {@code Sync} also
     * fans out to that sink.
     *
     * @param config connector config
     * @param files the authorized files API
     * @param attachmentMaxBytes server-wide inline byte limit; 0 = no limit
     * @param downstream optional extra sink for {@code Sync} emissions;
     *        {@code null} for none
     */
    public MicrosoftGrpcService(MicrosoftConnectorConfig config, GraphFiles files,
            long attachmentMaxBytes, MicrosoftChangeSink downstream) {
        this(config, files, attachmentMaxBytes, downstream, null, DEFAULT_DOWNLOAD_CHUNK_BYTES);
    }

    /**
     * Creates a service that can also resolve catalog connections.
     *
     * @param config default (environment) connector config
     * @param files the default connection's authorized files API
     * @param attachmentMaxBytes server-wide inline byte limit; 0 = no limit
     * @param downstream optional extra sink for {@code Sync}; {@code null} for none
     * @param connections {@code ConnectionService} stub for {@code connection_id}
     *        lookups; {@code null} when no catalog is configured
     * @param downloadChunkBytes default chunk size for {@code DownloadContent};
     *        must be positive
     */
    public MicrosoftGrpcService(MicrosoftConnectorConfig config, GraphFiles files,
            long attachmentMaxBytes, MicrosoftChangeSink downstream,
            ConnectionServiceGrpc.ConnectionServiceBlockingStub connections, int downloadChunkBytes) {
        Objects.requireNonNull(config, "config");
        Objects.requireNonNull(files, "files");
        if (attachmentMaxBytes < 0) {
            throw new IllegalArgumentException("attachmentMaxBytes must not be negative");
        }
        if (downloadChunkBytes <= 0) {
            throw new IllegalArgumentException("downloadChunkBytes must be positive");
        }
        this.sessions = new MicrosoftSessions(config, files, connections);
        this.mapper = new MicrosoftMapper();
        this.attachmentMaxBytes = attachmentMaxBytes;
        this.downstream = downstream;
        this.downloadChunkBytes = downloadChunkBytes;
    }

    // ------------------------------------------------------------ identity

    @Override
    public void getMe(GetMeRequest request, StreamObserver<GetMeResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            requireDelegated(session, "/me");
            JsonNode me = session.pager().get(List.of("/me?$select=" + GraphPaths.USER_SELECT, "/me"),
                    Map.of());
            return GetMeResponse.newBuilder().setUser(valid(mapper.toUser(me))).build();
        });
    }

    @Override
    public void probeConnection(ProbeConnectionRequest request,
            StreamObserver<ProbeConnectionResponse> observer) {
        unary(observer, () -> {
            boolean inline = !request.getTenantId().isBlank() || !request.getClientId().isBlank()
                    || !request.getClientSecret().isBlank();
            MicrosoftSessions.Session session;
            if (inline) {
                if (request.getTenantId().isBlank() || request.getClientId().isBlank()
                        || request.getClientSecret().isBlank()) {
                    throw new IllegalArgumentException(
                            "an inline probe needs tenant_id, client_id and client_secret");
                }
                session = sessions.inline(request.getTenantId(), request.getClientId(),
                        request.getClientSecret());
            } else {
                session = sessions.resolve(request.getConnectionId());
            }
            int limit = request.getLimit() <= 0 ? DEFAULT_PROBE_SITES : request.getLimit();
            ProbeConnectionResponse.Builder response = ProbeConnectionResponse.newBuilder()
                    .setConnectionId(inline ? "" : session.connectionId());
            try {
                AuthMode mode = session.authMode();
                String route = mode == AuthMode.AUTH_MODE_APPLICATION
                        ? "/sites/getAllSites?$select=id" : "/sites?search=*&$select=id";
                List<String> ids = new ArrayList<>();
                session.pager().forEach(List.of(route), Map.of(), node -> {
                    ids.add(node.path("id").asText(""));
                    return ids.size() < limit;
                });
                response.setOk(true).setAuthMode(session.authMode()).addAllSiteIds(ids);
                sessions.recordProbe(session.connectionId(), true, null);
            } catch (IOException | RuntimeException e) {
                String message = scrub(String.valueOf(e.getMessage()), request.getClientSecret());
                response.setOk(false).setAuthMode(safeMode(session)).setErrorMessage(message);
                if (!inline) {
                    sessions.recordProbe(session.connectionId(), false, message);
                }
            }
            return response.build();
        });
    }

    // ------------------------------------------------------------ sites

    @Override
    public void listSites(ListSitesRequest request, StreamObserver<ListSitesResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String query = request.getQuery().isBlank() ? "*" : request.getQuery();
            String base = "/sites?search=" + GraphPaths.query(query);
            List<Site> sites = new ArrayList<>();
            int cap = request.getLimit();
            session.pager().forEach(List.of(base + "&$select=" + GraphPaths.SITE_SELECT, base),
                    Map.of(), node -> {
                        sites.add(valid(mapper.toSite(node)));
                        return cap <= 0 || sites.size() < cap;
                    });
            return ListSitesResponse.newBuilder().addAllSites(sites).build();
        });
    }

    @Override
    public void listAllSites(ListAllSitesRequest request,
            StreamObserver<ListAllSitesResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            requireApplication(session, "/sites/getAllSites");
            String base = "/sites/getAllSites";
            session.pager().forEach(List.of(base + "?$select=" + GraphPaths.SITE_SELECT, base),
                    Map.of(), node -> {
                        Site site = valid(mapper.toSite(node));
                        if (request.getIncludePersonalSites() || !site.getIsPersonalSite()) {
                            observer.onNext(ListAllSitesResponse.newBuilder().setSite(site).build());
                        }
                        return true;
                    });
        });
    }

    @Override
    public void getSite(GetSiteRequest request, StreamObserver<GetSiteResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route = "/sites/" + GraphPaths.site(request.getSiteId());
            Site site = valid(mapper.toSite(session.pager().get(
                    List.of(route + "?$select=" + GraphPaths.SITE_SELECT, route), Map.of())));
            return GetSiteResponse.newBuilder().setSite(site).build();
        });
    }

    @Override
    public void listSubsites(ListSubsitesRequest request,
            StreamObserver<ListSubsitesResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route = "/sites/" + GraphPaths.site(request.getSiteId()) + "/sites";
            session.pager().forEach(List.of(route + "?$select=" + GraphPaths.SITE_SELECT, route),
                    Map.of(), node -> {
                        Site site = mapper.toSite(node);
                        if (!site.hasParentReference()) {
                            site = site.toBuilder().setParentReference(ItemReference.newBuilder()
                                    .setSiteId(request.getSiteId())).build();
                        }
                        observer.onNext(ListSubsitesResponse.newBuilder().setSite(valid(site)).build());
                        return true;
                    });
        });
    }

    @Override
    public void deltaSites(DeltaSitesRequest request, StreamObserver<DeltaSitesResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            requireApplication(session, "/sites/delta");
            String base = "/sites/delta";
            DeltaStart start = deltaStart(session, request.getDeltaLink(),
                    List.of(base + "?$select=" + GraphPaths.SITE_SELECT, base), base, "sites");
            GraphPager.DeltaOutcome outcome = session.pager().delta(start.fresh(), start.link(),
                    Map.of(), new GraphPager.DeltaVisitor() {
                        @Override
                        public void resync(DeltaResync resync) {
                            observer.onNext(DeltaSitesResponse.newBuilder().setResync(resync).build());
                        }

                        @Override
                        public void item(JsonNode node) {
                            observer.onNext(DeltaSitesResponse.newBuilder()
                                    .setSite(valid(mapper.toSite(node))).build());
                        }
                    });
            observer.onNext(DeltaSitesResponse.newBuilder().setDeltaLink(outcome.deltaLink()).build());
        });
    }

    // ------------------------------------------------------------ drives

    @Override
    public void listDrives(ListDrivesRequest request, StreamObserver<ListDrivesResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route;
            String siteId = "";
            if (!request.getSiteId().isBlank()) {
                route = "/sites/" + GraphPaths.site(request.getSiteId()) + "/drives";
                siteId = request.getSiteId();
            } else if (!request.getUserId().isBlank()) {
                route = "/users/" + GraphPaths.segment(request.getUserId(), "user_id") + "/drives";
            } else if (!request.getGroupId().isBlank()) {
                route = "/groups/" + GraphPaths.segment(request.getGroupId(), "group_id") + "/drives";
            } else {
                requireDelegated(session, "/me/drives");
                route = "/me/drives";
            }
            String owner = siteId;
            List<Drive> drives = new ArrayList<>();
            session.pager().forEach(List.of(route + "?$select=" + GraphPaths.DRIVE_SELECT, route),
                    Map.of(), node -> {
                        drives.add(valid(mapper.toDrive(node, owner)));
                        return true;
                    });
            return ListDrivesResponse.newBuilder().addAllDrives(drives).build();
        });
    }

    @Override
    public void getDrive(GetDriveRequest request, StreamObserver<GetDriveResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route = GraphPaths.drive(request.getDriveId());
            Drive drive = valid(mapper.toDrive(session.pager().get(
                    List.of(route + "?$select=" + GraphPaths.DRIVE_SELECT, route), Map.of())));
            return GetDriveResponse.newBuilder().setDrive(drive).build();
        });
    }

    @Override
    public void listChildren(ListChildrenRequest request,
            StreamObserver<ListChildrenResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            DriveItemExpander expander = expander(session, request.hasInclude(), request.getInclude(),
                    false);
            String route = request.getFolderId().isBlank()
                    ? childrenByPath(request.getDriveId(), request.getFolderPath())
                    : GraphPaths.item(request.getDriveId(), request.getFolderId()) + "/children";
            session.pager().forEach(expander.candidates(route), Map.of(), node -> {
                DriveItem item = expander.expand(mapper.toDriveItem(node, request.getDriveId()));
                observer.onNext(ListChildrenResponse.newBuilder().setItem(valid(item)).build());
                return true;
            });
        });
    }

    @Override
    public void getItem(GetItemRequest request, StreamObserver<GetItemResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            DriveItemExpander expander = expander(session, request.hasInclude(), request.getInclude(),
                    true);
            String route = request.getItemId().isBlank()
                    ? GraphPaths.itemByPath(request.getDriveId(), request.getPath())
                    : GraphPaths.item(request.getDriveId(), request.getItemId());
            DriveItem item = readItem(session, expander, route, request.getDriveId());
            return GetItemResponse.newBuilder().setItem(valid(item)).build();
        });
    }

    @Override
    @SuppressWarnings("deprecation")
    public void downloadItem(DownloadItemRequest request,
            StreamObserver<DownloadItemResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            DriveItemExpander expander = expander(session, request.hasInclude(), request.getInclude(),
                    true);
            DriveItem item = readItem(session, expander,
                    GraphPaths.item(request.getDriveId(), request.getItemId()), request.getDriveId());
            if (request.getIncludeContent()) {
                item = withContent(session, item, inlineLimit(request.getMaxInlineBytes()));
            }
            return DownloadItemResponse.newBuilder().setItem(valid(item)).build();
        });
    }

    @Override
    public void downloadContent(DownloadContentRequest request,
            StreamObserver<DownloadContentResponse> observer) {
        stream(observer, () -> {
            if (request.getOffset() < 0 || request.getLength() < 0 || request.getChunkBytes() < 0) {
                throw new IllegalArgumentException("offset, length and chunk_bytes must not be negative");
            }
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            DriveItemExpander expander = expander(session, request.hasInclude(), request.getInclude(),
                    true);
            String itemRoute = GraphPaths.item(request.getDriveId(), request.getItemId());
            DriveItem item = readItem(session, expander, itemRoute, request.getDriveId());
            if (item.getFolder() || item.hasPackage()) {
                throw Status.FAILED_PRECONDITION
                        .withDescription("item " + item.getId() + " is a folder; no binary")
                        .asRuntimeException();
            }
            FlowControl flow = FlowControl.of(observer);
            flow.awaitReady();
            observer.onNext(DownloadContentResponse.newBuilder().setItem(valid(item)).build());
            String route = request.getVersionId().isBlank() ? itemRoute + "/content"
                    : itemRoute + "/versions/" + GraphPaths.segment(request.getVersionId(), "version_id")
                            + "/content";
            if (!request.getFormat().isBlank()) {
                route += "?format=" + GraphPaths.query(request.getFormat());
            }
            long offset = request.getOffset();
            long length = request.getLength();
            String range = offset == 0 && length == 0 ? ""
                    : "bytes=" + offset + "-" + (length > 0 ? Long.toString(offset + length - 1) : "");
            int chunk = request.getChunkBytes() == 0 ? downloadChunkBytes : request.getChunkBytes();
            try (GraphClient.Content content = session.graph().openContent(route, range)) {
                InputStream in = content.body();
                if (content.status() != 206 && offset > 0) {
                    // Graph ignored the range: apply it here, with HTTP 416's rule
                    // (an offset at or past the end is not satisfiable).
                    try {
                        in.skipNBytes(offset);
                    } catch (java.io.EOFException e) {
                        throw outOfRange(item, offset);
                    }
                    int probe = in.read();
                    if (probe < 0) {
                        throw outOfRange(item, offset);
                    }
                    in = new java.io.SequenceInputStream(
                            new java.io.ByteArrayInputStream(new byte[] {(byte) probe}), in);
                }
                long remaining = length > 0 ? length : Long.MAX_VALUE;
                while (remaining > 0) {
                    byte[] buffer = in.readNBytes((int) Math.min(chunk, remaining));
                    if (buffer.length == 0) {
                        break;
                    }
                    remaining -= buffer.length;
                    flow.awaitReady();
                    observer.onNext(DownloadContentResponse.newBuilder()
                            .setChunk(ByteString.copyFrom(buffer)).build());
                }
            }
        });
    }

    private static StatusRuntimeException outOfRange(DriveItem item, long offset) {
        return Status.OUT_OF_RANGE
                .withDescription("offset " + offset + " is at or past the end of item " + item.getId())
                .asRuntimeException();
    }

    @Override
    public void deltaDriveItems(DeltaDriveItemsRequest request,
            StreamObserver<DeltaDriveItemsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String driveId = request.getDriveId();
            String drive = GraphPaths.drive(driveId);
            String folder = request.getItemId();
            String routeBase;
            String[] routeSegments;
            if (folder.isBlank() || folder.equalsIgnoreCase("root")) {
                routeBase = drive + "/root/delta";
                routeSegments = new String[] {"drives", driveId, "root"};
            } else {
                JsonNode driveNode = session.graph().get(drive + "?$select=id,driveType", Map.of());
                if (!"personal".equals(driveNode.path("driveType").asText(""))) {
                    throw new IllegalArgumentException("folder-scoped delta (item_id) is supported on "
                            + "OneDrive personal only; OneDrive for Business and SharePoint support "
                            + "delta on the drive root");
                }
                routeBase = GraphPaths.item(driveId, folder) + "/delta";
                routeSegments = new String[] {"drives", driveId, "items", folder};
            }
            DriveItemExpander expander = expander(session, request.hasInclude(), request.getInclude(),
                    false);
            Map<String, String> headers = DeltaHeaders.prefer(request.getPreferences());
            DeltaStart start = deltaStart(session, request.getDeltaLink(),
                    expander.candidates(routeBase), routeBase, routeSegments);
            GraphPager.DeltaOutcome outcome = session.pager().delta(start.fresh(), start.link(), headers,
                    new GraphPager.DeltaVisitor() {
                        @Override
                        public void resync(DeltaResync resync) {
                            observer.onNext(DeltaDriveItemsResponse.newBuilder().setResync(resync).build());
                        }

                        @Override
                        public void item(JsonNode node) throws IOException, InterruptedException {
                            DriveItem item = expander.expand(mapper.toDriveItem(node, driveId));
                            observer.onNext(DeltaDriveItemsResponse.newBuilder()
                                    .setItem(valid(item)).build());
                        }
                    });
            observer.onNext(DeltaDriveItemsResponse.newBuilder()
                    .setDeltaLink(outcome.deltaLink()).build());
        });
    }

    @Override
    public void listItemVersions(ListItemVersionsRequest request,
            StreamObserver<ListItemVersionsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route = GraphPaths.item(request.getDriveId(), request.getItemId()) + "/versions";
            session.pager().forEach(List.of(route), Map.of(), node -> {
                observer.onNext(ListItemVersionsResponse.newBuilder()
                        .setVersion(valid(mapper.toVersion(node, request.getDriveId(),
                                request.getItemId())))
                        .build());
                return true;
            });
        });
    }

    // ------------------------------------------------------------ lists

    @Override
    public void listLists(ListListsRequest request, StreamObserver<ListListsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String route = "/sites/" + GraphPaths.site(request.getSiteId()) + "/lists";
            session.pager().forEach(List.of(route + "?$select=" + GraphPaths.LIST_SELECT, route),
                    Map.of(), node -> {
                        SharePointList list = mapper.toList(node, request.getSiteId());
                        if (request.getIncludeHidden()
                                || (!list.getIsSystem() && !list.getList().getHidden())) {
                            observer.onNext(ListListsResponse.newBuilder().setList(valid(list)).build());
                        }
                        return true;
                    });
        });
    }

    @Override
    public void getList(GetListRequest request, StreamObserver<GetListResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            SharePointList list;
            if (!request.getListId().isBlank()) {
                String route = GraphPaths.list(request.getSiteId(), request.getListId());
                list = mapper.toList(session.pager().get(
                        List.of(route + "?$select=" + GraphPaths.LIST_SELECT, route), Map.of()),
                        request.getSiteId());
            } else if (!request.getDriveId().isBlank()) {
                String route = GraphPaths.drive(request.getDriveId()) + "/list";
                list = mapper.toList(session.pager().get(
                        List.of(route + "?$select=" + GraphPaths.LIST_SELECT, route), Map.of()),
                        request.getSiteId()).toBuilder().setDriveId(request.getDriveId()).build();
            } else {
                throw new IllegalArgumentException("site_id with list_id, or drive_id, is required");
            }
            return GetListResponse.newBuilder().setList(valid(list)).build();
        });
    }

    @Override
    public void listListItems(ListListItemsRequest request,
            StreamObserver<ListListItemsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            ListItemReader reader = reader(session, request.getSiteId(), request.getListId(),
                    request.getInclude());
            Map<String, String> headers = request.getFilter().isBlank() ? Map.of()
                    : Map.of("Prefer", GraphPaths.HONOR_NON_INDEXED);
            session.pager().forEach(reader.listCandidates(request.getFilter()), headers, node -> {
                observer.onNext(ListListItemsResponse.newBuilder()
                        .setItem(valid(reader.read(node))).build());
                return true;
            });
        });
    }

    @Override
    public void getListItem(GetListItemRequest request, StreamObserver<GetListItemResponse> observer) {
        unary(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            ListItemReader reader = reader(session, request.getSiteId(), request.getListId(),
                    request.getInclude());
            JsonNode node = session.pager().get(reader.itemCandidates(request.getItemId()), Map.of());
            return GetListItemResponse.newBuilder().setItem(valid(reader.read(node))).build();
        });
    }

    @Override
    public void deltaListItems(DeltaListItemsRequest request,
            StreamObserver<DeltaListItemsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String siteId = canonicalSiteId(session, request.getSiteId(), request.getDeltaLink());
            ListItemReader reader = reader(session, siteId, request.getListId(), request.getInclude());
            String base = GraphPaths.list(siteId, request.getListId()) + "/items/delta";
            DeltaStart start = deltaStart(session, request.getDeltaLink(), reader.deltaCandidates(), base,
                    "sites", siteId, "lists", request.getListId(), "items");
            GraphPager.DeltaOutcome outcome = session.pager().delta(start.fresh(), start.link(), Map.of(),
                    new GraphPager.DeltaVisitor() {
                        @Override
                        public void resync(DeltaResync resync) {
                            observer.onNext(DeltaListItemsResponse.newBuilder().setResync(resync).build());
                        }

                        @Override
                        public void item(JsonNode node) throws IOException, InterruptedException {
                            observer.onNext(DeltaListItemsResponse.newBuilder()
                                    .setItem(valid(reader.read(node))).build());
                        }
                    });
            observer.onNext(DeltaListItemsResponse.newBuilder().setDeltaLink(outcome.deltaLink()).build());
        });
    }

    @Override
    public void listColumns(ListColumnsRequest request, StreamObserver<ListColumnsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            String site = "/sites/" + GraphPaths.site(request.getSiteId());
            String route;
            String listId = "";
            String contentTypeId = "";
            if (!request.getListId().isBlank()) {
                route = GraphPaths.list(request.getSiteId(), request.getListId()) + "/columns";
                listId = request.getListId();
            } else if (!request.getContentTypeId().isBlank()) {
                route = site + "/contentTypes/"
                        + GraphPaths.segment(request.getContentTypeId(), "content_type_id") + "/columns";
                contentTypeId = request.getContentTypeId();
            } else {
                route = site + "/columns";
            }
            String list = listId;
            String contentType = contentTypeId;
            session.pager().forEach(List.of(route), Map.of(), node -> {
                ColumnDefinition column = mapper.toColumn(node, request.getSiteId(), list, contentType);
                observer.onNext(ListColumnsResponse.newBuilder().setColumn(valid(column)).build());
                return true;
            });
        });
    }

    // ------------------------------------------------------------ permissions

    @Override
    public void listPermissions(ListPermissionsRequest request,
            StreamObserver<ListPermissionsResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            PermissionTarget target = request.getTarget();
            String route = switch (target.getTargetCase()) {
                case DRIVE_ITEM -> GraphPaths.item(target.getDriveItem().getDriveId(),
                        target.getDriveItem().getItemId()) + "/permissions";
                case LIST_ITEM -> GraphPaths.listItem(target.getListItem().getSiteId(),
                        target.getListItem().getListId(), target.getListItem().getItemId())
                        + "/permissions";
                case LIST -> GraphPaths.list(target.getList().getSiteId(), target.getList().getListId())
                        + "/permissions";
                case SITE_ID -> "/sites/" + GraphPaths.site(target.getSiteId()) + "/permissions";
                case TARGET_NOT_SET -> throw new IllegalArgumentException("target is required");
            };
            session.pager().forEach(List.of(route), Map.of(), node -> {
                observer.onNext(ListPermissionsResponse.newBuilder()
                        .setPermission(valid(mapper.toPermission(node, target))).build());
                return true;
            });
        });
    }

    // ------------------------------------------------------------ sync

    @Override
    public void sync(SyncRequest request, StreamObserver<SyncResponse> observer) {
        stream(observer, () -> {
            MicrosoftSessions.Session session = sessions.resolve(request.getConnectionId());
            SyncCursor since = SyncCursors.decode(request.getSinceCursor(), session.graph().baseUrl());
            DriveItemExpander.Flags flags = DriveItemExpander.Flags.of(request.hasInclude(),
                    request.getInclude(), true);
            ListItemInclude listInclude = request.getListItemInclude();
            MicrosoftCrawler.Options options = new MicrosoftCrawler.Options(
                    request.getDriveIdsList(), request.getSiteIdsList(),
                    request.getFolderPath().isBlank() ? session.config().folderPath()
                            : request.getFolderPath(),
                    request.getIncludeContent(), inlineLimit(request.getMaxInlineBytes()),
                    flags.listColumns(), flags.permissions(), flags.retentionLabel(), flags.publication(),
                    request.getIncludeLists(), listInclude.getDriveItem(), listInclude.getPermissions(),
                    request.getDeltaPreferences(), request.getIncludeSystem(), since);
            DeltaHeaders.prefer(options.preferences());
            Object lock = new Object();
            MicrosoftChangeSink observerSink = new MicrosoftChangeSink() {
                @Override
                public void emit(MicrosoftChange change) {
                    synchronized (lock) {
                        observer.onNext(SyncResponse.newBuilder().setChange(change).build());
                    }
                }

                @Override
                public void snapshot(MicrosoftSnapshot snapshot) {
                    synchronized (lock) {
                        observer.onNext(SyncResponse.newBuilder().setSnapshot(snapshot).build());
                    }
                }
            };
            MicrosoftChangeSink sink = downstream == null ? observerSink
                    : new CompositeMicrosoftChangeSink(List.of(observerSink,
                            downstream.forConnection(session.connectionId())));
            MicrosoftCrawler.Result result = new MicrosoftCrawler(session.config(), session.files(), sink,
                    options).crawl();
            sink.completeRun(result.runId(), result.fullSweep());
            synchronized (lock) {
                observer.onNext(SyncResponse.newBuilder()
                        .setResumeCursor(SyncCursors.encode(result.cursor())).build());
            }
        });
    }

    // ------------------------------------------------------------ helpers

    private DriveItemExpander expander(MicrosoftSessions.Session session, boolean present,
            ai.pipestream.microsoft.v1.DriveItemInclude include, boolean defaultListColumns) {
        return new DriveItemExpander(session.files(), mapper,
                DriveItemExpander.Flags.of(present, include, defaultListColumns));
    }

    private ListItemReader reader(MicrosoftSessions.Session session, String siteId, String listId,
            ListItemInclude include) {
        GraphPaths.list(siteId, listId);
        return new ListItemReader(session.pager(), mapper, siteId, listId, include.getDriveItem(),
                include.getPermissions());
    }

    private DriveItem readItem(MicrosoftSessions.Session session, DriveItemExpander expander,
            String route, String driveId) throws IOException, InterruptedException {
        DriveItem item = mapper.toDriveItem(session.pager().get(expander.candidates(route), Map.of()),
                driveId);
        return expander.expand(item);
    }

    private static String childrenByPath(String driveId, String folderPath) {
        String encoded = GraphPaths.encodePath(folderPath);
        return encoded.isEmpty() ? GraphPaths.drive(driveId) + "/root/children"
                : GraphPaths.drive(driveId) + "/root:" + encoded + ":/children";
    }

    /** The request's inline limit, or the server-wide one when the request leaves 0. */
    private long inlineLimit(long requested) {
        if (requested < 0) {
            throw new IllegalArgumentException("max_inline_bytes must not be negative");
        }
        return requested > 0 ? requested : attachmentMaxBytes;
    }

    private DriveItem withContent(MicrosoftSessions.Session session, DriveItem item, long max)
            throws IOException, InterruptedException {
        if (item.getFolder() || item.hasPackage()) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " is a folder; no binary")
                    .asRuntimeException();
        }
        if (max > 0 && item.getSize() > max) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " is " + item.getSize()
                            + " bytes, above the " + max + " byte inline limit; use DownloadContent")
                    .asRuntimeException();
        }
        byte[] bytes = DriveItemExpander.inline(session.files(), item.getDriveId(), item.getId(), max);
        if (bytes == null) {
            throw Status.FAILED_PRECONDITION
                    .withDescription("item " + item.getId() + " downloaded to more than " + max
                            + " bytes, above the inline limit; use DownloadContent")
                    .asRuntimeException();
        }
        return item.toBuilder().setContent(ByteString.copyFrom(bytes)).build();
    }

    /**
     * Where a delta stream starts.
     *
     * @param fresh candidate URLs when no link is resumed
     * @param link validated client link, or empty
     */
    private record DeltaStart(List<String> fresh, String link) {
    }

    private static DeltaStart deltaStart(MicrosoftSessions.Session session, String deltaLink,
            List<String> fresh, String routeBase, String... routeSegments) {
        if (deltaLink == null || deltaLink.isBlank()) {
            return new DeltaStart(fresh, "");
        }
        if (deltaLink.trim().equalsIgnoreCase("latest")) {
            return new DeltaStart(List.of(routeBase + "?token=latest"), "");
        }
        GraphLinks.requireDeltaLink(session.graph().baseUrl(), deltaLink, routeSegments);
        return new DeltaStart(fresh, deltaLink);
    }

    /**
     * The composite site id for a {@code hostname:/path} site id when a delta link
     * has to be matched against it (Graph's links name the composite id).
     */
    private static String canonicalSiteId(MicrosoftSessions.Session session, String siteId,
            String deltaLink) throws IOException, InterruptedException {
        if (siteId == null || !siteId.contains(":/") || deltaLink == null || deltaLink.isBlank()
                || deltaLink.trim().equalsIgnoreCase("latest")) {
            return siteId;
        }
        return session.graph().get("/sites/" + GraphPaths.site(siteId) + "?$select=id", Map.of())
                .path("id").asText(siteId);
    }

    private static void requireDelegated(MicrosoftSessions.Session session, String route) {
        if (session.authMode() == AuthMode.AUTH_MODE_APPLICATION) {
            throw Status.FAILED_PRECONDITION
                    .withDescription(route + " needs a delegated (signed-in user) token; this "
                            + "connection holds an application token")
                    .asRuntimeException();
        }
    }

    private static void requireApplication(MicrosoftSessions.Session session, String route) {
        if (session.authMode() == AuthMode.AUTH_MODE_DELEGATED) {
            throw Status.FAILED_PRECONDITION
                    .withDescription(route + " needs an application (client-credentials) token; "
                            + "this connection holds a delegated token")
                    .asRuntimeException();
        }
    }

    private static AuthMode safeMode(MicrosoftSessions.Session session) {
        try {
            return session.authMode();
        } catch (RuntimeException e) {
            return AuthMode.AUTH_MODE_UNSPECIFIED;
        }
    }

    /** Removes {@code secret} from {@code message} (defence in depth: Entra never echoes it). */
    static String scrub(String message, String secret) {
        if (secret == null || secret.isEmpty() || message == null) {
            return message;
        }
        return message.replace(secret, "***");
    }

    private static <T extends Message> T valid(T message) {
        VALIDATOR.requireValid(message);
        return message;
    }

    @FunctionalInterface
    private interface UnaryBody<T> {
        T run() throws Exception;
    }

    @FunctionalInterface
    private interface StreamBody {
        void run() throws Exception;
    }

    private static <T> void unary(StreamObserver<T> observer, UnaryBody<T> body) {
        T response;
        try {
            response = body.run();
        } catch (Throwable t) {
            fail(observer, t);
            return;
        }
        observer.onNext(response);
        observer.onCompleted();
    }

    private static void stream(StreamObserver<?> observer, StreamBody body) {
        try {
            body.run();
        } catch (Throwable t) {
            fail(observer, t);
            return;
        }
        observer.onCompleted();
    }

    private static void fail(StreamObserver<?> observer, Throwable t) {
        Throwable cause = t;
        if (cause instanceof java.io.UncheckedIOException u && u.getCause() != null) {
            cause = u.getCause();
        }
        if (cause instanceof StatusRuntimeException e) {
            observer.onError(e);
            return;
        }
        Status status;
        if (cause instanceof GraphClient.GraphApiException e) {
            status = (switch (e.status()) {
                case 400 -> Status.INVALID_ARGUMENT;
                case 401 -> Status.UNAUTHENTICATED;
                case 403 -> Status.PERMISSION_DENIED;
                case 404 -> Status.NOT_FOUND;
                case 409, 412 -> Status.ABORTED;
                case 410 -> Status.FAILED_PRECONDITION;
                case 416 -> Status.OUT_OF_RANGE;
                case 429 -> Status.RESOURCE_EXHAUSTED;
                case 503, 504 -> Status.UNAVAILABLE;
                default -> Status.INTERNAL;
            }).withDescription(e.getMessage());
        } else if (cause instanceof GraphPager.RepeatedLinkException e) {
            status = Status.INTERNAL.withDescription(e.getMessage());
        } else if (cause instanceof IllegalArgumentException e) {
            status = Status.INVALID_ARGUMENT.withDescription(e.getMessage());
        } else if (cause instanceof InterruptedException e) {
            Thread.currentThread().interrupt();
            status = Status.CANCELLED.withDescription("interrupted: " + e.getMessage());
        } else if (cause instanceof IOException e) {
            status = Status.UNAVAILABLE.withDescription("Graph unreachable: " + e.getMessage());
        } else if (cause instanceof IllegalStateException e) {
            status = Status.UNAVAILABLE.withDescription(e.getMessage());
        } else {
            status = Status.INTERNAL.withDescription(String.valueOf(cause.getMessage()));
        }
        observer.onError(status.asRuntimeException());
    }

    /**
     * The download loop's cancellation check. Readiness itself is enforced
     * for every stream by {@link FlowControlInterceptor} (each send parks the
     * handler's virtual thread until the transport is ready, so a slow
     * client never makes the server buffer a whole file); this stops the
     * read loop as soon as the client has gone.
     */
    private static final class FlowControl {
        private final io.grpc.Context context = io.grpc.Context.current();

        static FlowControl of(StreamObserver<?> observer) {
            return new FlowControl();
        }

        void awaitReady() {
            if (context.isCancelled()) {
                throw Status.CANCELLED.withDescription("client cancelled the download")
                        .asRuntimeException();
            }
        }
    }
}
