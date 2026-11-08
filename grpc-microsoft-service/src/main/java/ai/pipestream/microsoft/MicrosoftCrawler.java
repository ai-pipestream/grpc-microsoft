package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DeltaPreferences;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.SharePointList;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.microsoft.v1.SyncCursor;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;

import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;
import java.util.UUID;

/**
 * One bounded, delta-driven sync pass over OneDrive / SharePoint. Every drive in
 * scope is swept with {@code /drives/{id}/root/delta}: a full enumeration when
 * the incoming {@link SyncCursor} has no link for it, otherwise only the changes
 * since that link, deletes included (the {@code deleted} facet becomes a
 * {@code CHANGE_OPERATION_DELETE}). With {@link Options#includeLists()} the
 * non-library lists of the sites in scope are swept the same way through
 * {@code /lists/{id}/items/delta}. Graph's {@code 410 Gone} restarts that
 * drive's or list's enumeration from scratch and its snapshot reports
 * {@code CHANGE_SOURCE_CRAWL}. Each sweep ends with one {@link MicrosoftSnapshot}
 * carrying its new delta link and whether it was a full enumeration;
 * {@link Result#cursor()} gathers every link for the next pass.
 */
public final class MicrosoftCrawler {

    /** Default cap for inlined file bytes: 25 MiB. */
    public static final long DEFAULT_ATTACHMENT_MAX_BYTES = 25L * 1024 * 1024;

    /** Count key for the drive root item a delta enumeration returns. */
    static final String COUNT_ROOT = "root";

    /**
     * What one pass covers and fetches.
     *
     * @param driveIds drives to sweep in addition to the drives of {@code siteIds}
     * @param siteIds sites whose drives (and, with {@code includeLists}, lists) are swept
     * @param folderPath non-root path: only items under it are emitted
     * @param includeContent inline file bytes up to {@code maxInlineBytes}
     * @param maxInlineBytes largest binary to inline; 0 = no limit
     * @param listColumns fetch each drive item's list-item columns
     * @param permissions fetch each item's permissions
     * @param retentionLabel fetch each drive item's retention label
     * @param publication select the publication facet
     * @param includeLists sweep non-library lists of the sites in scope
     * @param listItemDriveItem expand each list item's drive item view
     * @param listItemPermissions fetch each list item's permissions
     * @param preferences Graph {@code Prefer} preferences for drive sweeps
     * @param includeSystem also sweep system drives and hidden / system lists
     * @param since the previous pass's cursor; empty = full enumeration
     */
    public record Options(List<String> driveIds, List<String> siteIds, String folderPath,
            boolean includeContent, long maxInlineBytes, boolean listColumns, boolean permissions,
            boolean retentionLabel, boolean publication, boolean includeLists,
            boolean listItemDriveItem, boolean listItemPermissions, DeltaPreferences preferences,
            boolean includeSystem, SyncCursor since) {

        /** Copies the lists and defaults nulls. */
        public Options {
            driveIds = driveIds == null ? List.of() : List.copyOf(driveIds);
            siteIds = siteIds == null ? List.of() : List.copyOf(siteIds);
            folderPath = folderPath == null ? "" : folderPath;
            preferences = preferences == null ? DeltaPreferences.getDefaultInstance() : preferences;
            since = since == null ? SyncCursor.getDefaultInstance() : since;
        }

        /**
         * The first-revision defaults: list columns on, nothing else, no content.
         *
         * @param config the connection's scope (drive allowlist, folder path)
         * @param includeContent whether to inline bytes
         * @param maxInlineBytes inline byte limit; 0 = none
         * @return the options
         */
        public static Options defaults(MicrosoftConnectorConfig config, boolean includeContent,
                long maxInlineBytes) {
            return new Options(List.of(), List.of(), config.folderPath(), includeContent,
                    maxInlineBytes, true, false, false, false, false, false, false,
                    DeltaPreferences.getDefaultInstance(), false, SyncCursor.getDefaultInstance());
        }
    }

    /**
     * What a pass produced.
     *
     * @param runId the pass's run id (the cursor of its full-sweep changes)
     * @param fullSweep true when every sweep in the pass was a full enumeration,
     *        so a sink may delete what the pass did not emit
     * @param cursor every delta link for the next pass (links of drives and lists
     *        outside this pass's scope are carried over unchanged)
     */
    public record Result(String runId, boolean fullSweep, SyncCursor cursor) {
    }

    private final MicrosoftConnectorConfig config;
    private final GraphFiles files;
    private final GraphPager pager;
    private final MicrosoftMapper mapper;
    private final MicrosoftChangeSink sink;
    private final Options options;
    private final DriveItemExpander expander;
    private final MicrosoftValidator validator = MicrosoftValidator.create();

    /**
     * Creates a crawler with the first-revision defaults, without inlining bytes.
     *
     * @param config crawl scope and credentials
     * @param files the authorized files API
     * @param sink where changes and snapshots go
     */
    public MicrosoftCrawler(MicrosoftConnectorConfig config, GraphFiles files,
            MicrosoftChangeSink sink) {
        this(config, files, sink, Options.defaults(config, false, 0));
    }

    /**
     * Creates a crawler with the first-revision defaults. When
     * {@code includeContent} is set, file bytes up to {@code attachmentMaxBytes}
     * (0 = no limit) are inlined on each drive item.
     *
     * @param config crawl scope and credentials
     * @param files the authorized files API
     * @param sink where changes and snapshots go
     * @param attachmentMaxBytes inline file byte limit; 0 = none, never negative
     * @param includeContent whether to fetch file content
     */
    public MicrosoftCrawler(MicrosoftConnectorConfig config, GraphFiles files,
            MicrosoftChangeSink sink, long attachmentMaxBytes, boolean includeContent) {
        this(config, files, sink, Options.defaults(config, includeContent, attachmentMaxBytes));
        if (attachmentMaxBytes < 0) {
            throw new IllegalArgumentException("attachmentMaxBytes must not be negative");
        }
    }

    /**
     * Creates a crawler with explicit options.
     *
     * @param config connection scope (drive allowlist, site) used when the options
     *        name no drives or sites
     * @param files the authorized files API
     * @param sink where changes and snapshots go
     * @param options what the pass covers and fetches
     */
    public MicrosoftCrawler(MicrosoftConnectorConfig config, GraphFiles files,
            MicrosoftChangeSink sink, Options options) {
        this.config = Objects.requireNonNull(config, "config");
        this.files = Objects.requireNonNull(files, "files");
        this.pager = new GraphPager(files.client());
        this.mapper = new MicrosoftMapper();
        this.sink = Objects.requireNonNull(sink, "sink");
        this.options = Objects.requireNonNull(options, "options");
        this.expander = new DriveItemExpander(files, mapper, new DriveItemExpander.Flags(
                options.listColumns(), options.permissions(), options.retentionLabel(),
                options.publication()));
    }

    /**
     * Runs the pass.
     *
     * @return the run id, whether it was a full sweep, and the next cursor
     * @throws IOException if a Graph call fails
     * @throws InterruptedException if the calling thread is interrupted while waiting
     */
    public Result crawl() throws IOException, InterruptedException {
        String runId = UUID.randomUUID().toString();
        SyncCursor.Builder cursor = options.since().toBuilder();
        boolean incrementalPass = options.since().getDriveDeltaLinksCount() > 0
                || options.since().getListDeltaLinksCount() > 0;
        ChangeSource passSource = incrementalPass ? ChangeSource.CHANGE_SOURCE_DELTA
                : ChangeSource.CHANGE_SOURCE_CRAWL;
        boolean fullSweep = true;

        List<Site> sites = new ArrayList<>();
        for (String siteId : scopeSites()) {
            Site site = mapper.toSite(pager.get(List.of(
                    "/sites/" + GraphPaths.site(siteId) + "?$select=" + GraphPaths.SITE_SELECT,
                    "/sites/" + GraphPaths.site(siteId)), Map.of()));
            sites.add(site);
            emit(entity(site.getId()).setSite(site).build(), ChangeOperation.CHANGE_OPERATION_UPSERT,
                    passSource, runId);
        }
        for (Drive drive : scopeDrives(sites)) {
            fullSweep &= sweepDrive(drive, runId, passSource, cursor);
        }
        if (options.includeLists()) {
            for (Site site : sites) {
                fullSweep &= sweepLists(site, runId, passSource, cursor);
            }
        }
        return new Result(runId, fullSweep, cursor.build());
    }

    private List<String> scopeSites() {
        if (!options.siteIds().isEmpty()) {
            return options.siteIds();
        }
        if (options.driveIds().isEmpty() && !config.hasDriveAllowlist() && !config.siteId().isBlank()) {
            return List.of(config.siteId());
        }
        return List.of();
    }

    private List<Drive> scopeDrives(List<Site> sites) throws IOException, InterruptedException {
        Map<String, Drive> drives = new LinkedHashMap<>();
        List<String> explicit = options.driveIds().isEmpty() && options.siteIds().isEmpty()
                ? config.driveIds() : options.driveIds();
        for (String driveId : explicit) {
            Drive drive = mapper.toDrive(pager.get(List.of(
                    GraphPaths.drive(driveId) + "?$select=" + GraphPaths.DRIVE_SELECT,
                    GraphPaths.drive(driveId)), Map.of()));
            drives.putIfAbsent(drive.getId(), drive);
        }
        for (Site site : sites) {
            String base = "/sites/" + GraphPaths.site(site.getId()) + "/drives";
            pager.forEach(List.of(base + "?$select=" + GraphPaths.DRIVE_SELECT, base), Map.of(), node -> {
                Drive drive = mapper.toDrive(node, site.getId());
                if (options.includeSystem() || !drive.getIsSystem()) {
                    drives.putIfAbsent(drive.getId(), drive);
                }
                return true;
            });
        }
        if (drives.isEmpty() && explicit.isEmpty() && sites.isEmpty()) {
            Drive me = mapper.toDrive(files.meDrive());
            drives.put(me.getId(), me);
        }
        return new ArrayList<>(drives.values());
    }

    /** Returns whether the sweep was a full enumeration. */
    private boolean sweepDrive(Drive drive, String runId, ChangeSource passSource,
            SyncCursor.Builder cursor) throws IOException, InterruptedException {
        Instant started = Instant.now();
        Map<String, Long> counts = new TreeMap<>();
        emit(entity(drive.getId()).setDrive(drive).build(), ChangeOperation.CHANGE_OPERATION_UPSERT,
                passSource, runId);
        counts.merge("drive", 1L, Long::sum);
        String link = cursor.getDriveDeltaLinksOrDefault(drive.getId(), "");
        ChangeSource[] source = {link.isEmpty() ? ChangeSource.CHANGE_SOURCE_CRAWL
                : ChangeSource.CHANGE_SOURCE_DELTA};
        String changeCursor = link.isEmpty() ? runId : link;
        String[] readUnder = {changeCursor};
        FolderScope scope = FolderScope.of(pager, drive.getId(), options.folderPath());
        String base = GraphPaths.drive(drive.getId()) + "/root/delta";
        GraphPager.DeltaOutcome outcome = pager.delta(
                List.of(base + "?$select=" + GraphPaths.driveItemSelect(options.publication()), base),
                link, DeltaHeaders.prefer(options.preferences()), new GraphPager.DeltaVisitor() {
                    @Override
                    public void resync(ai.pipestream.microsoft.v1.DeltaResync resync) {
                        source[0] = ChangeSource.CHANGE_SOURCE_CRAWL;
                        readUnder[0] = runId;
                    }

                    @Override
                    public void item(JsonNode node) throws IOException, InterruptedException {
                        DriveItem item = mapper.toDriveItem(node, drive.getId());
                        if (item.getId().isEmpty()) {
                            return;
                        }
                        if (item.hasDeleted()) {
                            if (scope.containsDeleted(item)) {
                                emit(entity(item.getId()).setDriveItem(item).build(),
                                        ChangeOperation.CHANGE_OPERATION_DELETE, source[0], readUnder[0]);
                                counts.merge("deleted", 1L, Long::sum);
                            }
                            return;
                        }
                        if (!scope.contains(item)) {
                            return;
                        }
                        item = enrich(item);
                        emit(entity(item.getId()).setDriveItem(item).build(),
                                ChangeOperation.CHANGE_OPERATION_UPSERT, source[0], readUnder[0]);
                        counts.merge(item.getIsRoot() ? COUNT_ROOT : item.getFolder() ? "folder" : "file",
                                1L, Long::sum);
                    }
                });
        cursor.putDriveDeltaLinks(drive.getId(), outcome.deltaLink());
        snapshot(MicrosoftSnapshot.newBuilder()
                .setSnapshotId(runId + "-" + drive.getId())
                .setDriveId(drive.getId())
                .setSiteId(drive.getSiteId())
                .putAllEntityCounts(counts)
                .setCursor(outcome.deltaLink())
                .setStartedAt(timestamp(started))
                .setCompletedAt(timestamp(Instant.now()))
                .setSource(source[0])
                .build());
        return source[0] == ChangeSource.CHANGE_SOURCE_CRAWL;
    }

    private DriveItem enrich(DriveItem item) throws IOException, InterruptedException {
        item = expander.expand(item);
        if (options.includeContent() && !item.getFolder() && !item.hasPackage()
                && !item.hasMalware() && item.getSize() > 0
                && (options.maxInlineBytes() == 0 || item.getSize() <= options.maxInlineBytes())) {
            byte[] bytes = DriveItemExpander.inline(files, item.getDriveId(), item.getId(),
                    options.maxInlineBytes());
            if (bytes != null) {
                item = item.toBuilder().setContent(ByteString.copyFrom(bytes)).build();
            }
        }
        return item;
    }

    /** Returns whether every list sweep of the site was a full enumeration. */
    private boolean sweepLists(Site site, String runId, ChangeSource passSource,
            SyncCursor.Builder cursor) throws IOException, InterruptedException {
        List<SharePointList> lists = new ArrayList<>();
        String base = "/sites/" + GraphPaths.site(site.getId()) + "/lists";
        pager.forEach(List.of(base + "?$select=" + GraphPaths.LIST_SELECT, base), Map.of(), node -> {
            SharePointList list = mapper.toList(node, site.getId());
            boolean library = "documentLibrary".equals(list.getList().getTemplate());
            boolean hidden = list.getIsSystem() || list.getList().getHidden();
            if (!library && (options.includeSystem() || !hidden)) {
                lists.add(list);
            }
            return true;
        });
        boolean full = true;
        for (SharePointList list : lists) {
            full &= sweepList(site.getId(), list, runId, passSource, cursor);
        }
        return full;
    }

    private boolean sweepList(String siteId, SharePointList list, String runId,
            ChangeSource passSource, SyncCursor.Builder cursor) throws IOException, InterruptedException {
        Instant started = Instant.now();
        Map<String, Long> counts = new TreeMap<>();
        emit(entity(list.getId()).setList(list).build(), ChangeOperation.CHANGE_OPERATION_UPSERT,
                passSource, runId);
        counts.merge("list", 1L, Long::sum);
        String key = SyncCursors.listKey(siteId, list.getId());
        String link = cursor.getListDeltaLinksOrDefault(key, "");
        ChangeSource[] source = {link.isEmpty() ? ChangeSource.CHANGE_SOURCE_CRAWL
                : ChangeSource.CHANGE_SOURCE_DELTA};
        String[] readUnder = {link.isEmpty() ? runId : link};
        ListItemReader reader = new ListItemReader(pager, mapper, siteId, list.getId(),
                options.listItemDriveItem(), options.listItemPermissions());
        GraphPager.DeltaOutcome outcome = pager.delta(reader.deltaCandidates(), link, Map.of(),
                new GraphPager.DeltaVisitor() {
                    @Override
                    public void resync(ai.pipestream.microsoft.v1.DeltaResync resync) {
                        source[0] = ChangeSource.CHANGE_SOURCE_CRAWL;
                        readUnder[0] = runId;
                    }

                    @Override
                    public void item(JsonNode node) throws IOException, InterruptedException {
                        ListItem item = reader.read(node);
                        if (item.getId().isEmpty()) {
                            return;
                        }
                        boolean deleted = item.hasDeleted();
                        emit(entity(MicrosoftEntityIds.listItem(item)).setListItem(item).build(),
                                deleted ? ChangeOperation.CHANGE_OPERATION_DELETE
                                        : ChangeOperation.CHANGE_OPERATION_UPSERT,
                                source[0], readUnder[0]);
                        counts.merge(deleted ? "deleted" : "list_item", 1L, Long::sum);
                    }
                });
        cursor.putListDeltaLinks(key, outcome.deltaLink());
        snapshot(MicrosoftSnapshot.newBuilder()
                .setSnapshotId(runId + "-" + list.getId())
                .setSiteId(siteId)
                .setListId(list.getId())
                .putAllEntityCounts(counts)
                .setCursor(outcome.deltaLink())
                .setStartedAt(timestamp(started))
                .setCompletedAt(timestamp(Instant.now()))
                .setSource(source[0])
                .build());
        return source[0] == ChangeSource.CHANGE_SOURCE_CRAWL;
    }

    private void emit(MicrosoftEntity entity, ChangeOperation operation, ChangeSource source,
            String cursor) {
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId(UUID.randomUUID().toString())
                .setOperation(operation)
                .setEntity(entity)
                .setCursor(cursor)
                .setSource(source)
                .setOccurredAt(timestamp(Instant.now()))
                .build();
        validator.requireValid(change);
        sink.emit(change);
    }

    private void snapshot(MicrosoftSnapshot snapshot) {
        validator.requireValid(snapshot);
        sink.snapshot(snapshot);
    }

    private static MicrosoftEntity.Builder entity(String entityId) {
        return MicrosoftEntity.newBuilder()
                .setEntityId(entityId)
                .setIngestedAt(timestamp(Instant.now()));
    }

    private static Timestamp timestamp(Instant instant) {
        return Timestamp.newBuilder()
                .setSeconds(instant.getEpochSecond())
                .setNanos(instant.getNano())
                .build();
    }

    /**
     * Restricts a whole-drive delta feed to one folder subtree. Delta records carry
     * no parent path, so membership is decided by walking parent ids (each unknown
     * ancestor costs one {@code GET ?$select=id,parentReference}, cached).
     */
    static final class FolderScope {
        private final GraphPager pager;
        private final String driveId;
        private final String folderId;
        private final Map<String, Boolean> inside = new HashMap<>();

        private FolderScope(GraphPager pager, String driveId, String folderId) {
            this.pager = pager;
            this.driveId = driveId;
            this.folderId = folderId;
            if (folderId != null) {
                inside.put(folderId, Boolean.TRUE);
            }
        }

        static FolderScope of(GraphPager pager, String driveId, String folderPath)
                throws IOException, InterruptedException {
            if (GraphPaths.encodePath(folderPath).isEmpty()) {
                return new FolderScope(pager, driveId, null);
            }
            JsonNode folder = pager.get(List.of(GraphPaths.itemByPath(driveId, folderPath)
                    + "?$select=id"), Map.of());
            return new FolderScope(pager, driveId, folder.path("id").asText());
        }

        boolean contains(DriveItem item) throws IOException, InterruptedException {
            if (folderId == null) {
                return true;
            }
            if (item.getId().equals(folderId)) {
                return false;
            }
            boolean in = ancestorInside(item.getParentId());
            if (item.getFolder()) {
                inside.put(item.getId(), in);
            }
            return in;
        }

        boolean containsDeleted(DriveItem item) throws IOException, InterruptedException {
            if (folderId == null) {
                return true;
            }
            if (item.getParentId().isEmpty()) {
                return true;
            }
            try {
                return ancestorInside(item.getParentId());
            } catch (GraphClient.GraphApiException e) {
                if (e.status() == 404) {
                    return true;
                }
                throw e;
            }
        }

        private boolean ancestorInside(String parentId) throws IOException, InterruptedException {
            List<String> chain = new ArrayList<>();
            String current = parentId;
            Boolean verdict = null;
            while (current != null && !current.isEmpty()) {
                Boolean known = inside.get(current);
                if (known != null) {
                    verdict = known;
                    break;
                }
                chain.add(current);
                JsonNode node = pager.get(List.of(GraphPaths.item(driveId, current)
                        + "?$select=id,parentReference"), Map.of());
                current = node.path("parentReference").path("id").asText("");
            }
            boolean result = verdict != null && verdict;
            for (String id : chain) {
                inside.put(id, result);
            }
            return result;
        }
    }
}
