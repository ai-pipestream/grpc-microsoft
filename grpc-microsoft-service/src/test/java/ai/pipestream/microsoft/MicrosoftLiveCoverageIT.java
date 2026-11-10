package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.*;
import com.google.protobuf.Descriptors;
import com.google.protobuf.Message;
import io.grpc.Context;
import io.grpc.ManagedChannel;
import io.grpc.MethodDescriptor;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.URI;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.function.Predicate;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Calls every rpc of {@code MicrosoftService} against the real Microsoft 365
 * tenant and writes a coverage report ({@code build/live-coverage.md}, also
 * printed to stdout).
 *
 * <p>Gated on the same Entra <em>application</em> credentials as
 * {@link MicrosoftLiveSmokeIT} ({@code MICROSOFT_TENANT_ID},
 * {@code MICROSOFT_CLIENT_ID}, {@code MICROSOFT_CLIENT_SECRET}), optionally
 * {@code MICROSOFT_SITE_ID} to pin the site and {@code MICROSOFT_DRIVE_IDS}
 * to pin the drive. The token is app-only (client credentials), so
 * {@code GetMe} is expected to answer FAILED_PRECONDITION. Excluded from the
 * default {@code test} task; run with
 * {@code ./gradlew :grpc-microsoft-service:liveCoverageTest}.</p>
 *
 * <p>Ids are discovered first (site, drives, a folder, a file of at most
 * 10 MB, a list, a list item, a content type, ...); an entity kind the
 * tenant does not have is reported SKIPPED. Delta rpcs run one bounded full
 * round, then an incremental round from the delta link it returned. The
 * test fails only on defects in this codebase: INTERNAL / UNKNOWN that is
 * not an upstream 5xx, a validator rejection, a delta stream without its
 * delta link, a download byte mismatch, duplicate Sync entity ids, a
 * non-gRPC exception, or an rpc this test did not exercise. Permission and
 * not-found answers are reported, not failed. The client secret and the
 * access tokens never reach the report or stdout.</p>
 */
class MicrosoftLiveCoverageIT {

    private static final MicrosoftValidator VALIDATOR = MicrosoftValidator.create();
    private static final String TAG = "[MicrosoftLiveCoverageIT] ";

    /** Max items read from a list stream before the call is cancelled. */
    private static final int LIST_CAP = 5;
    /** Max candidate sites probed when picking the one with the most drives. */
    private static final int SITE_CANDIDATES = 5;
    /** Max children read per folder while looking for a folder and a file. */
    private static final int CHILDREN_CAP = 100;
    /** Max lists read from a site. */
    private static final int LISTS_CAP = 25;
    /** Max lists probed for an item. */
    private static final int LIST_ITEM_PROBES = 3;
    /** Item caps on the full delta rounds; past them the round is cancelled. */
    private static final int SITE_DELTA_CAP = 300;
    private static final int DRIVE_DELTA_CAP = 1000;
    private static final int LIST_DELTA_CAP = 500;
    /** Files above this size are not downloaded. */
    private static final long DOWNLOAD_CAP = 10L * 1024 * 1024;
    /** Length of the ranged DownloadContent probe. */
    private static final int RANGE_BYTES = 16;
    /** Safety valves on the one Sync pass (SyncRequest has no bound of its own). */
    private static final int SYNC_MESSAGE_CAP = 2000;
    private static final long SYNC_DEADLINE_MINUTES = 5;
    private static final long CALL_DEADLINE_SECONDS = 90;

    /** Rpcs whose upstream route exists for one auth mode only. */
    private static final Map<String, String> RPC_NOTES = Map.of(
            "GetMe", "delegated (signed-in user) token only; an app-only token has no /me");
    /** Calls that need more than Sites.Read.All + Files.Read.All upstream. */
    private static final Map<String, String> LABEL_NOTES = Map.of(
            "ListPermissions [site]", "needs Sites.FullControl.All",
            "ListPermissions [list item]", "may need Sites.FullControl.All");

    /** Office Open XML / legacy Office extensions: SharePoint rewrites their properties on download. */
    private static final Set<String> OFFICE_EXTENSIONS = Set.of(
            "docx", "docm", "dotx", "xlsx", "xlsm", "xltx", "pptx", "pptm", "potx",
            "doc", "xls", "ppt", "vsdx", "one");

    /** The Graph route each rpc is served by (the main one, for the report). */
    private static final Map<String, String> ROUTES = Map.ofEntries(
            Map.entry("GetMe", "GET /me"),
            Map.entry("ProbeConnection", "GET /sites/getAllSites"),
            Map.entry("ListSites", "GET /sites?search="),
            Map.entry("ListAllSites", "GET /sites/getAllSites"),
            Map.entry("GetSite", "GET /sites/{site-id}"),
            Map.entry("ListSubsites", "GET /sites/{site-id}/sites"),
            Map.entry("DeltaSites", "GET /sites/delta"),
            Map.entry("ListDrives", "GET /sites/{site-id}/drives"),
            Map.entry("GetDrive", "GET /drives/{drive-id}"),
            Map.entry("ListChildren", "GET /drives/{drive-id}/items/{id}/children"),
            Map.entry("GetItem", "GET /drives/{drive-id}/items/{id}"),
            Map.entry("DownloadItem", "GET /drives/{drive-id}/items/{id} + /content"),
            Map.entry("DownloadContent", "GET /drives/{drive-id}/items/{id}/content"),
            Map.entry("DeltaDriveItems", "GET /drives/{drive-id}/root/delta"),
            Map.entry("ListItemVersions", "GET /drives/{drive-id}/items/{id}/versions"),
            Map.entry("ListLists", "GET /sites/{site-id}/lists"),
            Map.entry("GetList", "GET /sites/{site-id}/lists/{list-id}, GET /drives/{drive-id}/list"),
            Map.entry("ListListItems", "GET /sites/{site-id}/lists/{list-id}/items"),
            Map.entry("GetListItem", "GET /sites/{site-id}/lists/{list-id}/items/{item-id}"),
            Map.entry("DeltaListItems", "GET /sites/{site-id}/lists/{list-id}/items/delta"),
            Map.entry("ListColumns", "GET /sites/{site-id}[/lists/{id}|/contentTypes/{id}]/columns"),
            Map.entry("ListPermissions", "GET .../permissions"),
            Map.entry("Sync", "delta crawl of one drive"));

    private record Row(String rpc, String label, String outcome, String detail, boolean failure) {
    }

    /** What one delta round streamed. */
    private record DeltaRound<T>(List<T> items, String link, boolean capped, int resyncs) {
    }

    /** One blocking call on the stub. */
    @FunctionalInterface
    private interface Call<T> {
        T apply(MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub);
    }

    private final List<Row> rows = new ArrayList<>();
    private final Map<String, String> discovered = new LinkedHashMap<>();
    private final List<String> deltaLines = new ArrayList<>();
    private final List<String> syncLines = new ArrayList<>();
    /** Credential values to scrub: the client secret and every access / refresh token issued. */
    private final Set<String> secrets = ConcurrentHashMap.newKeySet();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private MicrosoftConnectorConfig config;

    private static boolean notBlank(String value) {
        return value != null && !value.isBlank();
    }

    @Test
    void everyRpcAgainstTheLiveTenant() throws Exception {
        assumeTrue(notBlank(System.getenv(MicrosoftConnectorConfig.ENV_TENANT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_ID))
                        && notBlank(System.getenv(MicrosoftConnectorConfig.ENV_CLIENT_SECRET)),
                "no live Microsoft application credentials in the environment; skipping");
        Path out = Path.of(System.getProperty("liveCoverage.report", "build/live-coverage.md"));
        List<String> failures = run(MicrosoftConnectorConfig.fromEnvironment(), out);
        assertThat(failures).as("defects found by the live coverage run (see %s)", out).isEmpty();
    }

    /**
     * The whole run against the tenant of {@code config}: exercises every rpc,
     * writes and prints the report, and answers the failure lines (empty when
     * clean).
     */
    List<String> run(MicrosoftConnectorConfig config, Path out) throws Exception {
        this.config = config;
        secrets.add(config.clientSecret());
        GraphAuth auth = new GraphAuth(config.authConfig());
        AtomicReference<GraphAuth.Token> token = new AtomicReference<>(issue(auth));
        GraphClient client = new GraphClient(config.graphBaseUrl(), () -> {
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

        // A direct executor deadlocks with FlowControlInterceptor; handlers park virtual threads.
        ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();
        String name = InProcessServerBuilder.generateName();
        Server server = InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, new GraphFiles(client),
                        MicrosoftGrpcService.DEFAULT_ATTACHMENT_MAX_BYTES))
                .build().start();
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        Instant started = Instant.now();
        try {
            stub = MicrosoftServiceGrpc.newBlockingStub(channel);
            exercise();
        } finally {
            channel.shutdownNow();
            server.shutdownNow();
            server.awaitTermination(10, TimeUnit.SECONDS);
            executor.shutdownNow();
        }

        // Every rpc of the service must have been called (or skipped for lack of data).
        Set<String> seen = new LinkedHashSet<>();
        rows.forEach(row -> seen.add(row.rpc()));
        for (MethodDescriptor<?, ?> method : MicrosoftServiceGrpc.getServiceDescriptor().getMethods()) {
            String rpc = method.getBareMethodName();
            if (!seen.contains(rpc)) {
                rows.add(new Row(rpc, rpc, "FAIL", "not exercised by this test", true));
            }
        }

        String report = report(URI.create(config.graphBaseUrl()).getHost(), started);
        Files.createDirectories(out.toAbsolutePath().getParent());
        Files.writeString(out, report, StandardCharsets.UTF_8);
        System.out.println(report);
        System.out.println(TAG + "report written to " + out.toAbsolutePath());

        List<String> failures = new ArrayList<>();
        rows.stream().filter(Row::failure)
                .forEach(row -> failures.add(scrub(row.label() + ": " + row.outcome() + " " + row.detail())));
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

    // ======================================================================
    // THE RUN
    // ======================================================================

    private void exercise() {
        // ------------------------------------------------ identity / probe --
        unary("ProbeConnection", "ProbeConnection [configured credentials]",
                s -> s.probeConnection(ProbeConnectionRequest.newBuilder().setLimit(3).build()),
                MicrosoftLiveCoverageIT::probeDetail);
        unary("ProbeConnection", "ProbeConnection [inline credentials]",
                s -> s.probeConnection(ProbeConnectionRequest.newBuilder().setTenantId(config.tenantId())
                        .setClientId(config.clientId()).setClientSecret(config.clientSecret())
                        .setLimit(1).build()),
                MicrosoftLiveCoverageIT::probeDetail);
        unary("GetMe", "GetMe [app-only token]",
                s -> s.getMe(GetMeRequest.getDefaultInstance()),
                r -> "user " + r.getUser().getId());

        // ------------------------------------------------------- sites --
        Map<String, Site> candidates = new LinkedHashMap<>();
        ListSitesResponse search = unary("ListSites", "ListSites [search *]",
                s -> s.listSites(ListSitesRequest.newBuilder().setLimit(SITE_CANDIDATES).build()),
                r -> r.getSitesCount() + " site(s): " + join(r.getSitesList(), MicrosoftLiveCoverageIT::siteName));
        List<ListAllSitesResponse> all = stream("ListAllSites", "ListAllSites [no personal sites]",
                SITE_CANDIDATES, s -> s.listAllSites(ListAllSitesRequest.getDefaultInstance()),
                l -> join(l, r -> siteName(r.getSite())));
        String pinnedSite = config.siteId();
        if (notBlank(pinnedSite)) {
            candidates.put(pinnedSite, Site.newBuilder().setId(pinnedSite).build());
        } else {
            if (search != null) {
                search.getSitesList().forEach(site -> candidates.putIfAbsent(site.getId(), site));
            }
            if (all != null) {
                all.forEach(r -> candidates.putIfAbsent(r.getSite().getId(), r.getSite()));
            }
        }
        if (candidates.isEmpty()) {
            skipAll("no site visible to the credentials");
            return;
        }
        String siteId = null;
        int best = -1;
        int probed = 0;
        for (String id : candidates.keySet()) {
            if (probed++ >= SITE_CANDIDATES) {
                break;
            }
            int count = notBlank(pinnedSite) ? 0 : quietDriveCount(id);
            if (count > best) {
                best = count;
                siteId = id;
            }
        }
        String chosenSite = siteId;
        Site fetched = unary("GetSite", "GetSite [id]",
                s -> s.getSite(GetSiteRequest.newBuilder().setSiteId(chosenSite).build()),
                r -> siteName(r.getSite()) + ", " + r.getSite().getWebUrl(), GetSiteResponse::getSite);
        Site site = fetched != null ? fetched : candidates.get(chosenSite);
        discovered.put("site", siteName(site) + " (" + chosenSite + ")"
                + (notBlank(pinnedSite) ? ", from " + MicrosoftConnectorConfig.ENV_SITE_ID
                : ", " + best + " drive(s), the most among " + Math.min(candidates.size(), SITE_CANDIDATES)
                        + " candidates"));
        String pathRef = sitePathRef(site);
        if (pathRef == null) {
            discovered.put("site path form", "none (root site or no web_url)");
        } else {
            unary("GetSite", "GetSite [hostname:/path]",
                    s -> s.getSite(GetSiteRequest.newBuilder().setSiteId(pathRef).build()),
                    r -> "id " + r.getSite().getId());
        }
        stream("ListSubsites", "ListSubsites [site]", LIST_CAP,
                s -> s.listSubsites(ListSubsitesRequest.newBuilder().setSiteId(chosenSite).build()),
                l -> join(l, r -> siteName(r.getSite())));
        delta("DeltaSites", "tenant", SITE_DELTA_CAP,
                link -> s -> s.deltaSites(DeltaSitesRequest.newBuilder().setDeltaLink(link).build()),
                r -> r.getEventCase() == DeltaSitesResponse.EventCase.DELTA_LINK ? r.getDeltaLink() : null,
                r -> r.getEventCase() == DeltaSitesResponse.EventCase.RESYNC);

        // ------------------------------------------------------ drives --
        ListDrivesResponse drives = unary("ListDrives", "ListDrives [site]",
                s -> s.listDrives(ListDrivesRequest.newBuilder().setSiteId(chosenSite).build()),
                r -> r.getDrivesCount() + " drive(s): " + join(r.getDrivesList(),
                        d -> d.getName() + " (" + d.getDriveType() + ")"));
        Drive drive = null;
        Drive syncDrive = null;
        if (!config.driveIds().isEmpty()) {
            String pinned = config.driveIds().get(0);
            drive = Drive.newBuilder().setId(pinned).setName(pinned).build();
            syncDrive = drive;
            discovered.put("drive", pinned + ", from " + MicrosoftConnectorConfig.ENV_DRIVE_IDS);
        } else if (drives != null) {
            long smallest = Long.MAX_VALUE;
            for (Drive candidate : drives.getDrivesList()) {
                if (candidate.getIsSystem()) {
                    continue;
                }
                if (drive == null || "documentLibrary".equals(candidate.getDriveType())
                        && !"documentLibrary".equals(drive.getDriveType())) {
                    drive = candidate;
                }
                long used = candidate.getQuota().getUsed();
                if (used > 0 && used < smallest) {
                    smallest = used;
                    syncDrive = candidate;
                }
            }
            if (drive == null && drives.getDrivesCount() > 0) {
                drive = drives.getDrives(0);
            }
            if (syncDrive == null) {
                syncDrive = drive;
            }
        }
        if (drive == null) {
            for (String rpc : List.of("GetDrive", "ListChildren", "GetItem", "DownloadItem", "DownloadContent",
                    "DeltaDriveItems", "ListItemVersions", "Sync")) {
                skip(rpc, "no drive in the site");
            }
        } else {
            if (!discovered.containsKey("drive")) {
                discovered.put("drive (probes)", drive.getName() + " (" + drive.getId() + ") "
                        + drive.getDriveType());
                discovered.put("drive (Sync)", syncDrive.getName() + " (" + syncDrive.getId() + "), "
                        + syncDrive.getQuota().getUsed() + " bytes used");
            }
            String driveId = drive.getId();
            unary("GetDrive", "GetDrive",
                    s -> s.getDrive(GetDriveRequest.newBuilder().setDriveId(driveId).build()),
                    r -> r.getDrive().getName() + ", " + r.getDrive().getDriveType());
            unary("GetList", "GetList [drive's list]",
                    s -> s.getList(GetListRequest.newBuilder().setSiteId(chosenSite).setDriveId(driveId).build()),
                    r -> "list " + r.getList().getId() + " \"" + r.getList().getDisplayName() + "\"");
            driveItems(driveId);
        }

        // ------------------------------------------------------- lists --
        lists(chosenSite);

        // ------------------------------------------------- site ACL --
        stream("ListPermissions", "ListPermissions [site]", LIST_CAP,
                s -> s.listPermissions(ListPermissionsRequest.newBuilder()
                        .setTarget(PermissionTarget.newBuilder().setSiteId(chosenSite)).build()),
                l -> join(l, r -> r.getPermission().getId()));

        // -------------------------------------------------------- sync --
        if (syncDrive != null) {
            sync(syncDrive);
        }
    }

    private static String probeDetail(ProbeConnectionResponse r) {
        return "ok=" + r.getOk() + ", " + r.getAuthMode().name() + ", " + r.getSiteIdsCount() + " site id(s)"
                + (r.getOk() ? "" : ", " + r.getErrorMessage());
    }

    /** Drives of a site, for candidate ranking only; the real call below reports errors. */
    private int quietDriveCount(String siteId) {
        try {
            return stub().listDrives(ListDrivesRequest.newBuilder().setSiteId(siteId).build()).getDrivesCount();
        } catch (StatusRuntimeException e) {
            return -1;
        }
    }

    /**
     * Folders, a small file, its downloads, versions and ACL, and the drive
     * delta rounds.
     */
    private void driveItems(String driveId) {
        Map<DriveItem, String> paths = new LinkedHashMap<>();
        List<ListChildrenResponse> root = stream("ListChildren", "ListChildren [root]", CHILDREN_CAP,
                s -> s.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId).build()),
                l -> l.stream().filter(r -> r.getItem().getFolder()).count() + " folder(s)");
        DriveItem folder = null;
        if (root != null) {
            for (ListChildrenResponse row : root) {
                DriveItem item = row.getItem();
                paths.put(item, "/" + item.getName());
                if (folder == null && item.getFolder() && !item.hasPackage() && !item.hasRemoteItem()) {
                    folder = item;
                }
            }
        }
        if (folder == null) {
            discovered.put("folder", "none at the drive root");
        } else {
            DriveItem parent = folder;
            discovered.put("folder", parent.getId() + " \"" + parent.getName() + "\"");
            List<ListChildrenResponse> inside = stream("ListChildren", "ListChildren [folder id]", CHILDREN_CAP,
                    s -> s.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId)
                            .setFolderId(parent.getId()).build()),
                    l -> "");
            if (inside != null) {
                inside.forEach(r -> paths.put(r.getItem(), "/" + parent.getName() + "/" + r.getItem().getName()));
            }
            stream("ListChildren", "ListChildren [folder path]", LIST_CAP,
                    s -> s.listChildren(ListChildrenRequest.newBuilder().setDriveId(driveId)
                            .setFolderPath("/" + parent.getName()).build()),
                    l -> join(l, r -> r.getItem().getName()));
        }

        List<DeltaDriveItemsResponse> deltaItems = delta("DeltaDriveItems", "drive", DRIVE_DELTA_CAP,
                link -> s -> s.deltaDriveItems(DeltaDriveItemsRequest.newBuilder().setDriveId(driveId)
                        .setDeltaLink(link).build()),
                r -> r.getEventCase() == DeltaDriveItemsResponse.EventCase.DELTA_LINK ? r.getDeltaLink() : null,
                r -> r.getEventCase() == DeltaDriveItemsResponse.EventCase.RESYNC);

        DriveItem file = pickFile(new ArrayList<>(paths.keySet()));
        if (file == null) {
            List<DriveItem> fromDelta = new ArrayList<>();
            deltaItems.forEach(r -> fromDelta.add(r.getItem()));
            file = pickFile(fromDelta);
            if (file != null) {
                String path = drivePath(file);
                if (path != null) {
                    paths.put(file, path);
                }
            }
        }
        DriveItem target = file != null ? file : folder;
        if (target == null) {
            skip("GetItem", "no drive item found");
            skip("ListPermissions", "ListPermissions [drive item]", "no drive item found");
        } else {
            String itemId = target.getId();
            unary("GetItem", "GetItem [id, list columns + permissions + publication]",
                    s -> s.getItem(GetItemRequest.newBuilder().setDriveId(driveId).setItemId(itemId)
                            .setInclude(DriveItemInclude.newBuilder().setListColumns(true).setPermissions(true)
                                    .setPublication(true))
                            .build()),
                    r -> "\"" + r.getItem().getName() + "\", " + r.getItem().getSize() + " bytes, "
                            + r.getItem().getListColumnsCount() + " list columns, "
                            + r.getItem().getPermissionsCount() + " permissions");
            String path = paths.get(target);
            if (path != null) {
                unary("GetItem", "GetItem [path]",
                        s -> s.getItem(GetItemRequest.newBuilder().setDriveId(driveId).setPath(path).build()),
                        r -> "id " + r.getItem().getId() + (r.getItem().getId().equals(itemId) ? " (same item)"
                                : " (DIFFERENT item)"));
            }
            unary("GetItem", "GetItem [retention label]",
                    s -> s.getItem(GetItemRequest.newBuilder().setDriveId(driveId).setItemId(itemId)
                            .setInclude(DriveItemInclude.newBuilder().setRetentionLabel(true)).build()),
                    r -> r.getItem().hasRetentionLabel() ? "label \"" + r.getItem().getRetentionLabel().getName()
                            + "\"" : "no retention label");
            stream("ListPermissions", "ListPermissions [drive item]", LIST_CAP,
                    s -> s.listPermissions(ListPermissionsRequest.newBuilder().setTarget(PermissionTarget
                            .newBuilder().setDriveItem(DriveItemRef.newBuilder().setDriveId(driveId)
                                    .setItemId(itemId))).build()),
                    l -> join(l, r -> join(r.getPermission().getRolesList(), x -> x)));
        }
        if (file == null) {
            String reason = "no file of at most " + DOWNLOAD_CAP + " bytes among " + paths.size()
                    + " listed and " + deltaItems.size() + " delta items";
            skip("ListItemVersions", reason);
            skip("DownloadContent", reason);
            skip("DownloadItem", reason);
            return;
        }
        files(driveId, file);
    }

    /** Versions and the three downloads of one small file. */
    private void files(String driveId, DriveItem file) {
        String itemId = file.getId();
        boolean office = office(file.getName());
        discovered.put("file", itemId + " \"" + file.getName() + "\" " + file.getMimeType() + ", "
                + file.getSize() + " bytes" + (office ? " (Office document)" : ""));
        stream("ListItemVersions", "ListItemVersions [file]", LIST_CAP,
                s -> s.listItemVersions(ListItemVersionsRequest.newBuilder().setDriveId(driveId)
                        .setItemId(itemId).build()),
                l -> join(l, r -> r.getVersion().getId()));
        DriveItem meta = unary("GetItem", "GetItem [file, metadata only]",
                s -> s.getItem(GetItemRequest.newBuilder().setDriveId(driveId).setItemId(itemId)
                        .setInclude(DriveItemInclude.getDefaultInstance()).build()),
                r -> r.getItem().getSize() + " bytes", GetItemResponse::getItem);
        long size = meta == null ? file.getSize() : meta.getSize();
        DriveItemInclude bare = DriveItemInclude.getDefaultInstance();
        download("DownloadContent [full]", DownloadContentRequest.newBuilder().setDriveId(driveId)
                .setItemId(itemId).setInclude(bare).build(), size, true, office);
        download("DownloadContent [range 0-" + (RANGE_BYTES - 1) + "]", DownloadContentRequest.newBuilder()
                .setDriveId(driveId).setItemId(itemId).setLength(RANGE_BYTES).setInclude(bare).build(),
                Math.min(RANGE_BYTES, size), false, false);
        downloadItem(driveId, itemId, size, office);
    }

    /**
     * Reads a DownloadContent stream to the end and checks the byte count
     * against {@code expected} (and, when {@code compareHead}, against the
     * size the leading item message declares).
     */
    private void download(String label, DownloadContentRequest request, long expected, boolean compareHead,
            boolean office) {
        String rpc = "DownloadContent";
        Context.CancellableContext ctx = Context.current().withCancellation();
        Context previous = ctx.attach();
        try {
            Iterator<DownloadContentResponse> it = stub().downloadContent(request);
            long bytes = 0;
            long head = -1;
            int messages = 0;
            boolean dataBeforeHead = false;
            while (it.hasNext()) {
                DownloadContentResponse message = it.next();
                messages++;
                if (message.getPayloadCase() == DownloadContentResponse.PayloadCase.ITEM) {
                    if (!violations(rpc, label, message)) {
                        return;
                    }
                    head = message.getItem().getSize();
                } else {
                    if (head < 0) {
                        dataBeforeHead = true;
                    }
                    bytes += message.getChunk().size();
                }
            }
            List<String> problems = new ArrayList<>();
            if (head < 0 || dataBeforeHead) {
                problems.add("the item message did not come first");
            }
            if (compareHead && head >= 0 && head != bytes) {
                problems.add("item declared " + head + " bytes, streamed " + bytes);
            }
            if (expected != bytes) {
                problems.add("expected " + expected + " bytes, streamed " + bytes);
            }
            sized(rpc, label, bytes + " bytes in " + messages + " messages", problems, office);
        } catch (Throwable t) {
            error(rpc, label, t);
        } finally {
            ctx.detach(previous);
            ctx.cancel(null);
        }
    }

    // DownloadItem is deprecated in the contract, and still served: it is covered too.
    @SuppressWarnings("deprecation")
    private void downloadItem(String driveId, String itemId, long expected, boolean office) {
        String rpc = "DownloadItem";
        String label = "DownloadItem [inline content, deprecated]";
        try {
            DownloadItemResponse response = stub().downloadItem(DownloadItemRequest.newBuilder()
                    .setDriveId(driveId).setItemId(itemId).setIncludeContent(true)
                    .setMaxInlineBytes(DOWNLOAD_CAP).setInclude(DriveItemInclude.getDefaultInstance()).build());
            if (!violations(rpc, label, response)) {
                return;
            }
            long bytes = response.getItem().getContent().size();
            List<String> problems = new ArrayList<>();
            if (!response.getItem().hasContent()) {
                problems.add("no content inlined");
            }
            if (response.getItem().getSize() != bytes) {
                problems.add("item declared " + response.getItem().getSize() + " bytes, inlined " + bytes);
            }
            if (expected != bytes) {
                problems.add("expected " + expected + " bytes, inlined " + bytes);
            }
            sized(rpc, label, bytes + " bytes inlined", problems, office);
        } catch (Throwable t) {
            error(rpc, label, t);
        }
    }

    /** OK, or BYTE_MISMATCH (a failure unless SharePoint rewrote an Office document). */
    private void sized(String rpc, String label, String detail, List<String> problems, boolean office) {
        if (problems.isEmpty()) {
            rows.add(new Row(rpc, label, "OK", detail, false));
        } else if (office) {
            rows.add(new Row(rpc, label, "SIZE_DIFFERS_UPSTREAM", detail + "; " + String.join("; ", problems)
                    + " (Office document: SharePoint rewrites its properties on download, so Graph's size"
                    + " need not match the bytes; reported, not failed)", false));
        } else {
            rows.add(new Row(rpc, label, "BYTE_MISMATCH", detail + "; " + String.join("; ", problems), true));
        }
    }

    /** A plain file of at most {@link #DOWNLOAD_CAP} bytes, non-Office preferred. */
    private static DriveItem pickFile(List<DriveItem> items) {
        DriveItem chosen = null;
        for (DriveItem item : items) {
            if (item.getFolder() || item.hasPackage() || item.hasRemoteItem() || item.hasDeleted()
                    || item.getIsRoot() || item.getSize() <= 0 || item.getSize() > DOWNLOAD_CAP) {
                continue;
            }
            if (chosen == null || office(chosen.getName()) && !office(item.getName())) {
                chosen = item;
            }
        }
        return chosen;
    }

    private static boolean office(String name) {
        int dot = name.lastIndexOf('.');
        return dot >= 0 && OFFICE_EXTENSIONS.contains(name.substring(dot + 1).toLowerCase(Locale.ROOT));
    }

    /** The drive-relative path of an item from its parent reference ({@code /drives/x/root:/a}), or null. */
    private static String drivePath(DriveItem item) {
        String parent = item.getParentReference().getPath();
        int root = parent.indexOf("root:");
        if (root < 0) {
            return null;
        }
        return parent.substring(root + "root:".length()) + "/" + item.getName();
    }

    /** {@code hostname:/server-relative-path} from a site's web URL, or null for a root site. */
    private static String sitePathRef(Site site) {
        if (site == null || site.getWebUrl().isBlank()) {
            return null;
        }
        try {
            URI uri = URI.create(site.getWebUrl());
            String path = uri.getPath();
            if (uri.getHost() == null || path == null || path.isEmpty() || path.equals("/")) {
                return null;
            }
            return uri.getHost() + ":" + path;
        } catch (IllegalArgumentException e) {
            return null;
        }
    }

    /** Lists, list items, columns, list ACLs and the list delta rounds. */
    private void lists(String siteId) {
        List<ListListsResponse> lists = stream("ListLists", "ListLists [site]", LISTS_CAP,
                s -> s.listLists(ListListsRequest.newBuilder().setSiteId(siteId).build()),
                l -> join(l, r -> r.getList().getDisplayName() + " (" + r.getList().getList().getTemplate() + ")"));
        stream("ListColumns", "ListColumns [site]", LIST_CAP,
                s -> s.listColumns(ListColumnsRequest.newBuilder().setSiteId(siteId).build()),
                l -> join(l, r -> r.getColumn().getName()));
        if (lists == null || lists.isEmpty()) {
            for (String rpc : List.of("ListListItems", "GetListItem", "DeltaListItems")) {
                skip(rpc, "no list in the site");
            }
            skip("GetList", "GetList [site, list id]", "no list in the site");
            skip("ListPermissions", "ListPermissions [list]", "no list in the site");
            return;
        }
        // Plain lists first, document libraries after.
        List<SharePointList> ordered = new ArrayList<>();
        lists.forEach(r -> {
            if (!"documentLibrary".equals(r.getList().getList().getTemplate())) {
                ordered.add(r.getList());
            }
        });
        lists.forEach(r -> {
            if ("documentLibrary".equals(r.getList().getList().getTemplate())) {
                ordered.add(r.getList());
            }
        });
        SharePointList list = ordered.get(0);
        ListItem item = null;
        for (int i = 0; i < ordered.size() && i < LIST_ITEM_PROBES; i++) {
            SharePointList candidate = ordered.get(i);
            List<ListListItemsResponse> items = stream("ListListItems",
                    "ListListItems [list \"" + candidate.getDisplayName() + "\"]", LIST_CAP,
                    s -> s.listListItems(ListListItemsRequest.newBuilder().setSiteId(siteId)
                            .setListId(candidate.getId()).build()),
                    l -> join(l, r -> r.getItem().getId() + " (" + r.getItem().getFieldsCount() + " fields)"));
            if (items != null && !items.isEmpty()) {
                list = candidate;
                item = items.get(0).getItem();
                break;
            }
        }
        String listId = list.getId();
        discovered.put("list", listId + " \"" + list.getDisplayName() + "\" (" + list.getList().getTemplate() + ")");
        unary("GetList", "GetList [site, list id]",
                s -> s.getList(GetListRequest.newBuilder().setSiteId(siteId).setListId(listId).build()),
                r -> "\"" + r.getList().getDisplayName() + "\"");
        stream("ListColumns", "ListColumns [list]", LIST_CAP,
                s -> s.listColumns(ListColumnsRequest.newBuilder().setSiteId(siteId).setListId(listId).build()),
                l -> join(l, r -> r.getColumn().getName()));
        stream("ListPermissions", "ListPermissions [list]", LIST_CAP,
                s -> s.listPermissions(ListPermissionsRequest.newBuilder().setTarget(PermissionTarget.newBuilder()
                        .setList(ListRef.newBuilder().setSiteId(siteId).setListId(listId))).build()),
                l -> join(l, r -> r.getPermission().getId()));
        if (item == null) {
            skip("GetListItem", "no item in the first " + Math.min(ordered.size(), LIST_ITEM_PROBES) + " list(s)");
            skip("ListPermissions", "ListPermissions [list item]", "no list item found");
        } else {
            String itemId = item.getId();
            discovered.put("list item", listId + "/" + itemId);
            unary("GetListItem", "GetListItem",
                    s -> s.getListItem(GetListItemRequest.newBuilder().setSiteId(siteId).setListId(listId)
                            .setItemId(itemId).build()),
                    r -> r.getItem().getFieldsCount() + " fields, content type \""
                            + r.getItem().getContentType().getName() + "\"");
            stream("ListPermissions", "ListPermissions [list item]", LIST_CAP,
                    s -> s.listPermissions(ListPermissionsRequest.newBuilder().setTarget(PermissionTarget
                            .newBuilder().setListItem(ListItemRef.newBuilder().setSiteId(siteId)
                                    .setListId(listId).setItemId(itemId))).build()),
                    l -> join(l, r -> r.getPermission().getId()));
            String contentType = item.getContentType().getId();
            if (contentType.isBlank()) {
                discovered.put("content type", "none on the list item");
            } else {
                discovered.put("content type", contentType + " \"" + item.getContentType().getName() + "\"");
                stream("ListColumns", "ListColumns [content type]", LIST_CAP,
                        s -> s.listColumns(ListColumnsRequest.newBuilder().setSiteId(siteId)
                                .setContentTypeId(contentType).build()),
                        l -> join(l, r -> r.getColumn().getName()));
            }
        }
        delta("DeltaListItems", "list", LIST_DELTA_CAP,
                link -> s -> s.deltaListItems(DeltaListItemsRequest.newBuilder().setSiteId(siteId)
                        .setListId(listId).setDeltaLink(link).build()),
                r -> r.getEventCase() == DeltaListItemsResponse.EventCase.DELTA_LINK ? r.getDeltaLink() : null,
                r -> r.getEventCase() == DeltaListItemsResponse.EventCase.RESYNC);
    }

    /**
     * One bounded full delta round (falling back to {@code token=latest}
     * when the cap cancels it before its delta link), then one incremental
     * round from the link it returned. Answers the full round's messages.
     */
    private <T extends Message> List<T> delta(String rpc, String scope, int cap,
            Function<String, Call<Iterator<T>>> call, Function<T, String> linkOf, Predicate<T> isResync) {
        DeltaRound<T> full = deltaRound(rpc, rpc + " [" + scope + ", full]", cap, call.apply(""), linkOf,
                isResync);
        if (full == null) {
            deltaLines.add(rpc + " [" + scope + "]: full round failed; no incremental round");
            return List.of();
        }
        String link = full.link();
        String source = "the full round";
        if (full.capped()) {
            DeltaRound<T> latest = deltaRound(rpc, rpc + " [" + scope + ", token=latest]", cap,
                    call.apply("latest"), linkOf, isResync);
            link = latest == null ? "" : latest.link();
            source = "token=latest (the full round hit the " + cap + "-item cap)";
        }
        if (link.isEmpty()) {
            deltaLines.add(rpc + " [" + scope + "]: full round " + full.items().size() + " item(s); no delta link"
                    + " to resume from, no incremental round");
            return full.items();
        }
        DeltaRound<T> incremental = deltaRound(rpc, rpc + " [" + scope + ", incremental]", cap,
                call.apply(link), linkOf, isResync);
        deltaLines.add(rpc + " [" + scope + "]: full round " + full.items().size() + " item(s)"
                + (full.capped() ? " (capped)" : "") + ", delta link from " + source + "; incremental round "
                + (incremental == null ? "failed" : incremental.items().size() + " change(s)"
                + (incremental.resyncs() > 0 ? ", " + incremental.resyncs() + " resync" : "")
                + (incremental.link().isEmpty() ? "" : ", new delta link received")));
        return full.items();
    }

    /** One delta stream: its items (up to {@code cap}), its terminal delta link, and its resyncs. */
    private <T extends Message> DeltaRound<T> deltaRound(String rpc, String label, int cap, Call<Iterator<T>> call,
            Function<T, String> linkOf, Predicate<T> isResync) {
        Context.CancellableContext ctx = Context.current().withCancellation();
        List<T> items = new ArrayList<>();
        String[] link = {""};
        boolean[] sawLink = {false};
        boolean[] capped = {false};
        int[] resyncs = {0};
        try {
            ctx.run(() -> {
                Iterator<T> it = call.apply(stub());
                while (it.hasNext()) {
                    T message = it.next();
                    String value = linkOf.apply(message);
                    if (value != null) {
                        sawLink[0] = true;
                        link[0] = value;
                    } else if (isResync.test(message)) {
                        resyncs[0]++;
                    } else if (items.size() >= cap) {
                        capped[0] = true;
                        break;
                    } else {
                        items.add(message);
                    }
                }
            });
            for (T item : items) {
                if (!violations(rpc, label, item)) {
                    return null;
                }
            }
            String detail = items.size() + (capped[0] ? "+ (capped, cancelled before the delta link)" : "")
                    + " item(s)" + (resyncs[0] > 0 ? ", " + resyncs[0] + " resync" : "");
            if (capped[0]) {
                rows.add(new Row(rpc, label, "OK", detail, false));
            } else if (!sawLink[0] || link[0].isEmpty()) {
                rows.add(new Row(rpc, label, "NO_DELTA_LINK", detail + "; the stream ended without a delta link",
                        true));
            } else {
                rows.add(new Row(rpc, label, "OK", detail + ", delta link received (" + link[0].length()
                        + " chars)", false));
            }
            return new DeltaRound<>(items, link[0], capped[0], resyncs[0]);
        } catch (Throwable t) {
            error(rpc, label, t);
            return null;
        } finally {
            ctx.cancel(null);
        }
    }

    /**
     * One bounded Sync pass over one drive (no bodies): counts by kind,
     * snapshot counts, and duplicate entity ids (a failure).
     */
    private void sync(Drive drive) {
        String label = "Sync [one drive, no bodies]";
        Map<String, Long> counts = new TreeMap<>();
        Map<String, Long> operations = new TreeMap<>();
        Map<String, MicrosoftEntity> byId = new LinkedHashMap<>();
        List<String> duplicates = new ArrayList<>();
        Map<String, Long> snapshotCounts = new TreeMap<>();
        int[] snapshots = {0};
        int[] inlinePermissions = {0};
        boolean[] resumed = {false};
        int[] messages = {0};
        Context.CancellableContext ctx = Context.current().withCancellation();
        try {
            ctx.run(() -> {
                Iterator<SyncResponse> it = stub.withDeadlineAfter(SYNC_DEADLINE_MINUTES, TimeUnit.MINUTES)
                        .sync(SyncRequest.newBuilder().addDriveIds(drive.getId()).build());
                while (messages[0] < SYNC_MESSAGE_CAP && it.hasNext()) {
                    SyncResponse response = it.next();
                    messages[0]++;
                    List<MicrosoftValidator.Violation> v = validate(response);
                    if (!v.isEmpty()) {
                        rows.add(new Row("Sync", "Sync [validation, message " + messages[0] + "]", "MAPPING",
                                clip(v.toString()), true));
                    }
                    switch (response.getEventCase()) {
                        case CHANGE -> {
                            MicrosoftChange change = response.getChange();
                            operations.merge(change.getOperation().name(), 1L, Long::sum);
                            MicrosoftEntity entity = change.getEntity();
                            String kind = entity.getEntityCase().name().toLowerCase(Locale.ROOT);
                            counts.merge(kind, 1L, Long::sum);
                            if (entity.getDriveItem().getPermissionsCount() > 0
                                    || entity.getListItem().getPermissionsCount() > 0) {
                                inlinePermissions[0]++;
                            }
                            MicrosoftEntity previous = byId.putIfAbsent(entity.getEntityId(), entity);
                            if (previous != null) {
                                boolean same = previous.toBuilder().clearIngestedAt().build()
                                        .equals(entity.toBuilder().clearIngestedAt().build());
                                duplicates.add(entity.getEntityId() + " ("
                                        + previous.getEntityCase().name().toLowerCase(Locale.ROOT) + " then " + kind
                                        + (same ? ", identical payload" : ", different payload") + ")");
                            }
                        }
                        case SNAPSHOT -> {
                            snapshots[0]++;
                            response.getSnapshot().getEntityCountsMap()
                                    .forEach((k, n) -> snapshotCounts.merge(k, n, Long::sum));
                        }
                        case RESUME_CURSOR -> resumed[0] = true;
                        default -> {
                        }
                    }
                }
            });
            String detail = "drive " + drive.getName() + ": " + messages[0] + " messages, " + byId.size()
                    + " distinct entity ids" + (resumed[0] ? ", resume cursor received"
                    : ", TRUNCATED at the " + SYNC_MESSAGE_CAP + "-message cap (no resume cursor)");
            if (duplicates.isEmpty()) {
                rows.add(new Row("Sync", label, "OK", detail, false));
            } else {
                rows.add(new Row("Sync", label, "DUPLICATE_IDS",
                        detail + "; " + duplicates.size() + " duplicate entity ids", true));
            }
        } catch (Throwable t) {
            error("Sync", label, t);
        } finally {
            ctx.cancel(null);
        }
        syncLines.add("Drive: " + drive.getName() + " (" + drive.getId() + "), " + messages[0] + " messages, "
                + snapshots[0] + " snapshot(s), " + (resumed[0] ? "resume cursor received." : "no resume cursor."));
        syncLines.add("");
        syncLines.add("Entity counts by kind (from the change stream): " + (counts.isEmpty() ? "none" : ""));
        counts.forEach((k, n) -> syncLines.add("- " + k + ": " + n));
        syncLines.add("");
        syncLines.add("Change operations: " + (operations.isEmpty() ? "none" : ""));
        operations.forEach((k, n) -> syncLines.add("- " + k + ": " + n));
        syncLines.add("");
        syncLines.add("Snapshot entity_counts: " + (snapshotCounts.isEmpty() ? "none" : ""));
        snapshotCounts.forEach((k, n) -> syncLines.add("- " + k + ": " + n));
        syncLines.add("");
        syncLines.add("Permission completeness: not modeled (MicrosoftEntity has no completeness marker; ACLs"
                + " ride inline as DriveItem / ListItem permissions only when include.permissions is set,"
                + " which this bounded pass leaves off to keep request volume low). Entities carrying"
                + " inline permissions: " + inlinePermissions[0] + ".");
        syncLines.add("");
        syncLines.add("Duplicate entity_ids (failure): " + (duplicates.isEmpty() ? "none" : duplicates.size()));
        duplicates.stream().limit(50).forEach(line -> syncLines.add("- " + line));
        if (duplicates.size() > 50) {
            syncLines.add("- ... " + (duplicates.size() - 50) + " more");
        }
    }

    // ======================================================================
    // CALL HELPERS AND OUTCOME CLASSIFICATION
    // ======================================================================

    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub() {
        return stub.withDeadlineAfter(CALL_DEADLINE_SECONDS, TimeUnit.SECONDS);
    }

    /** One unary call; the response, or null when it failed (a row records why). */
    private <T extends Message, R> R unary(String rpc, String label, Call<T> call, Function<T, String> detail,
            Function<T, R> result) {
        try {
            T response = call.apply(stub());
            if (!violations(rpc, label, response)) {
                return null;
            }
            rows.add(new Row(rpc, label, "OK", clip(detail.apply(response)), false));
            return result.apply(response);
        } catch (Throwable t) {
            error(rpc, label, t);
            return null;
        }
    }

    private <T extends Message> T unary(String rpc, String label, Call<T> call, Function<T, String> detail) {
        return unary(rpc, label, call, detail, Function.identity());
    }

    private <T extends Message> List<T> stream(String rpc, String label, int max, Call<Iterator<T>> call,
            Function<List<T>, String> detail) {
        Context.CancellableContext ctx = Context.current().withCancellation();
        List<T> out = new ArrayList<>();
        try {
            ctx.run(() -> {
                Iterator<T> it = call.apply(stub());
                while (out.size() < max && it.hasNext()) {
                    out.add(it.next());
                }
            });
            for (T item : out) {
                if (!violations(rpc, label, item)) {
                    return null;
                }
            }
            String text = detail.apply(out);
            rows.add(new Row(rpc, label, "OK", out.size() + (out.size() >= max ? "+ (capped)" : "")
                    + " item(s)" + (text.isEmpty() ? "" : ": " + clip(text)), false));
            return out;
        } catch (Throwable t) {
            error(rpc, label, t);
            return null;
        } finally {
            ctx.cancel(null);
        }
    }

    /** Client-side re-validation of what the server sent; false (and a failure row) when invalid. */
    private boolean violations(String rpc, String label, Message message) {
        List<MicrosoftValidator.Violation> v = validate(message);
        if (v.isEmpty()) {
            return true;
        }
        rows.add(new Row(rpc, label, "MAPPING", clip(v.toString()), true));
        return false;
    }

    /**
     * The validator checks domain messages and ignores the rpc wrappers, so
     * validate the wrapper and each domain message it carries.
     */
    private static List<MicrosoftValidator.Violation> validate(Message message) {
        List<MicrosoftValidator.Violation> out = new ArrayList<>(VALIDATOR.validate(message).violations());
        message.getAllFields().forEach((field, value) -> {
            if (field.getJavaType() != Descriptors.FieldDescriptor.JavaType.MESSAGE || field.isMapField()) {
                return;
            }
            if (field.isRepeated()) {
                for (Object item : (List<?>) value) {
                    out.addAll(VALIDATOR.validate((Message) item).violations());
                }
            } else {
                out.addAll(VALIDATOR.validate((Message) value).violations());
            }
        });
        return out;
    }

    private void error(String rpc, String label, Throwable t) {
        if (!(t instanceof StatusRuntimeException e)) {
            rows.add(new Row(rpc, label, "EXCEPTION", clip(t.getClass().getName() + ": " + t.getMessage()), true));
            return;
        }
        Status status = e.getStatus();
        String description = status.getDescription() == null ? "" : status.getDescription();
        int http = httpStatus(description);
        String detail = ROUTES.getOrDefault(rpc, "?") + (http > 0 ? ", HTTP " + http : "") + ": " + description;
        String needs = LABEL_NOTES.getOrDefault(label, RPC_NOTES.get(rpc));
        String note = needs == null ? "" : " [" + needs + "]";
        switch (status.getCode()) {
            case NOT_FOUND, PERMISSION_DENIED, UNAUTHENTICATED, UNIMPLEMENTED, INVALID_ARGUMENT ->
                    rows.add(new Row(rpc, label, status.getCode().name(), clip(detail) + note, false));
            case INTERNAL, UNKNOWN -> {
                if (http >= 500) {
                    rows.add(new Row(rpc, label, "UPSTREAM_" + http, clip(detail) + note, false));
                } else if (http > 0 && !note.isEmpty()) {
                    rows.add(new Row(rpc, label, status.getCode().name(), clip(detail) + note
                            + " (reported, not failed)", false));
                } else {
                    rows.add(new Row(rpc, label, status.getCode().name(), clip(detail), true));
                }
            }
            case FAILED_PRECONDITION -> {
                // The server's own validator rejected what the mapper produced.
                boolean mapping = description.startsWith("microsoft validation failed");
                rows.add(new Row(rpc, label, mapping ? "MAPPING" : "FAILED_PRECONDITION", clip(detail) + note,
                        mapping));
            }
            default -> rows.add(new Row(rpc, label, status.getCode().name(), clip(detail) + note, false));
        }
    }

    /** The HTTP status in a "Graph 403 (accessDenied): ..." description, or 0. */
    private static int httpStatus(String description) {
        String prefix = "Graph ";
        int at = description.indexOf(prefix);
        if (at < 0) {
            return 0;
        }
        int value = 0;
        int digits = 0;
        for (int i = at + prefix.length(); i < description.length() && digits < 3; i++) {
            char c = description.charAt(i);
            if (c < '0' || c > '9') {
                break;
            }
            value = value * 10 + (c - '0');
            digits++;
        }
        return digits == 3 ? value : 0;
    }

    private void skip(String rpc, String reason) {
        skip(rpc, rpc, reason);
    }

    private void skip(String rpc, String label, String reason) {
        rows.add(new Row(rpc, label, "SKIPPED", "no data: " + reason, false));
    }

    private void skipAll(String reason) {
        for (MethodDescriptor<?, ?> method : MicrosoftServiceGrpc.getServiceDescriptor().getMethods()) {
            String rpc = method.getBareMethodName();
            if (rows.stream().noneMatch(row -> row.rpc().equals(rpc))) {
                skip(rpc, reason);
            }
        }
    }

    private static String siteName(Site site) {
        return site.getDisplayName().isBlank() ? site.getName() : site.getDisplayName();
    }

    private static <T> String join(List<T> list, Function<T, String> f) {
        List<String> parts = new ArrayList<>();
        list.forEach(item -> parts.add(f.apply(item)));
        return String.join(", ", parts);
    }

    private static String clip(String text) {
        String flat = text == null ? "" : text.replace('\n', ' ').replace('\r', ' ');
        return flat.length() <= 300 ? flat : flat.substring(0, 300) + "...";
    }

    // ======================================================================
    // REPORT
    // ======================================================================

    private String report(String host, Instant started) {
        StringBuilder md = new StringBuilder();
        md.append("# Microsoft Graph live coverage\n\n");
        md.append("- Graph: ").append(host).append(" (application token, client credentials)\n");
        md.append("- Run: ").append(started).append(" (")
                .append(Duration.between(started, Instant.now()).toSeconds()).append(" s)\n");
        long failures = rows.stream().filter(Row::failure).count();
        Map<String, Integer> outcomes = new TreeMap<>();
        rows.forEach(row -> outcomes.merge(row.outcome(), 1, Integer::sum));
        md.append("- Calls: ").append(rows.size()).append(", outcomes ").append(outcomes)
                .append(", failures ").append(failures).append("\n\n");

        md.append("## Discovered ids\n\n| entity | id / notes |\n|---|---|\n");
        discovered.forEach((k, v) -> md.append("| ").append(cell(k)).append(" | ").append(cell(v)).append(" |\n"));

        md.append("\n## RPCs\n\n| rpc | outcome | detail |\n|---|---|---|\n");
        for (Row row : rows) {
            md.append("| ").append(cell(row.label())).append(" | ")
                    .append(row.failure() ? "**" + row.outcome() + " (FAIL)**" : row.outcome())
                    .append(" | ").append(cell(row.detail())).append(" |\n");
        }

        md.append("\n## Delta rounds\n\n");
        if (deltaLines.isEmpty()) {
            md.append("none\n");
        }
        deltaLines.forEach(line -> md.append("- ").append(cell(line)).append('\n'));

        md.append("\n## Sync\n\n");
        if (syncLines.isEmpty()) {
            md.append("not run\n");
        }
        syncLines.forEach(line -> md.append(cell(line)).append('\n'));

        md.append("\n## Needs other credentials or permissions\n\n");
        md.append("The run holds an app-only token with Sites.Read.All + Files.Read.All in mind. These calls"
                + " are reported, never failed, when upstream refuses them:\n\n");
        new TreeMap<>(RPC_NOTES).forEach((rpc, why) -> md.append("- ").append(rpc).append(": ").append(why)
                .append('\n'));
        new TreeMap<>(LABEL_NOTES).forEach((label, why) -> md.append("- ").append(label).append(": ").append(why)
                .append('\n'));
        List<Row> refused = new ArrayList<>();
        rows.stream().filter(row -> row.outcome().equals("PERMISSION_DENIED")
                || row.outcome().equals("UNAUTHENTICATED")).forEach(refused::add);
        md.append("\nRefused upstream in this run: ").append(refused.isEmpty() ? "none" : "").append('\n');
        refused.forEach(row -> md.append("- ").append(cell(row.label())).append('\n'));

        md.append("\n## Failures\n\n");
        if (failures == 0) {
            md.append("none\n");
        } else {
            rows.stream().filter(Row::failure).forEach(row -> md.append("- ").append(cell(row.label()))
                    .append(": ").append(row.outcome()).append(' ').append(cell(row.detail())).append('\n'));
        }
        return scrub(md.toString());
    }

    private static String cell(String text) {
        return clip(text).replace("|", "\\|");
    }

    /** Belt and braces: the credential values never reach the report. */
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
