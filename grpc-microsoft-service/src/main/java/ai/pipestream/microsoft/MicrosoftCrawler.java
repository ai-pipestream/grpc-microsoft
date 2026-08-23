package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import com.fasterxml.jackson.databind.JsonNode;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;

import java.io.IOException;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;
import java.util.UUID;

/**
 * Walks OneDrive / SharePoint drives through {@link GraphFiles} and emits
 * {@link MicrosoftChange} UPSERTs plus one {@link MicrosoftSnapshot} per
 * drive. Folders are walked recursively; file bytes are optional and
 * size-capped.
 */
public final class MicrosoftCrawler {

    /** Default cap for inlined file bytes: 25 MiB. */
    public static final long DEFAULT_ATTACHMENT_MAX_BYTES = 25L * 1024 * 1024;

    private final MicrosoftConnectorConfig config;
    private final GraphFiles files;
    private final MicrosoftMapper mapper;
    private final MicrosoftChangeSink sink;
    private final long attachmentMaxBytes;
    private final boolean includeContent;

    /**
     * Creates a crawler that maps Graph entities and emits them to {@code sink},
     * without inlining file bytes.
     *
     * @param config crawl scope and credentials
     * @param files the authorized files API
     * @param sink where changes and snapshots go
     */
    public MicrosoftCrawler(MicrosoftConnectorConfig config, GraphFiles files,
            MicrosoftChangeSink sink) {
        this(config, files, sink, DEFAULT_ATTACHMENT_MAX_BYTES, false);
    }

    /**
     * Creates a crawler. When {@code includeContent} is {@code true}, file
     * bytes up to {@code attachmentMaxBytes} are inlined on each drive item.
     *
     * @param config crawl scope and credentials
     * @param files the authorized files API
     * @param sink where changes and snapshots go
     * @param attachmentMaxBytes inline file byte cap; must be positive
     * @param includeContent whether to fetch file content
     */
    public MicrosoftCrawler(MicrosoftConnectorConfig config, GraphFiles files,
            MicrosoftChangeSink sink, long attachmentMaxBytes, boolean includeContent) {
        this.config = Objects.requireNonNull(config, "config");
        this.files = Objects.requireNonNull(files, "files");
        this.mapper = new MicrosoftMapper();
        this.sink = Objects.requireNonNull(sink, "sink");
        if (attachmentMaxBytes <= 0) {
            throw new IllegalArgumentException("attachmentMaxBytes must be positive");
        }
        this.attachmentMaxBytes = attachmentMaxBytes;
        this.includeContent = includeContent;
    }

    /**
     * A full crawl of the configured drives (or the signed-in user's drive).
     *
     * @throws IOException if a Graph call fails
     * @throws InterruptedException if the calling thread is interrupted while waiting
     */
    public void crawl() throws IOException, InterruptedException {
        String runId = UUID.randomUUID().toString();
        for (Drive drive : listDrives()) {
            Instant started = Instant.now();
            Map<String, Long> counts = new TreeMap<>();
            emit(entity(drive.getId()).setDrive(drive).build(), ChangeSource.CHANGE_SOURCE_CRAWL,
                    runId);
            counts.merge("drive", 1L, Long::sum);
            crawlFolder(drive.getId(), config.folderPath(), ChangeSource.CHANGE_SOURCE_CRAWL,
                    runId, counts);
            MicrosoftSnapshot snapshot = MicrosoftSnapshot.newBuilder()
                    .setSnapshotId(runId + "-" + drive.getId())
                    .setDriveId(drive.getId())
                    .putAllEntityCounts(counts)
                    .setCursor(Instant.now().toString())
                    .setStartedAt(timestamp(started))
                    .setCompletedAt(timestamp(Instant.now()))
                    .build();
            MicrosoftValidator.create().requireValid(snapshot);
            sink.snapshot(snapshot);
        }
    }

    private List<Drive> listDrives() throws IOException, InterruptedException {
        List<Drive> drives = new ArrayList<>();
        if (config.hasDriveAllowlist()) {
            for (String driveId : config.driveIds()) {
                JsonNode node = files.drive(driveId);
                Drive drive = mapper.toDrive(node);
                drives.add(drive);
            }
            return drives;
        }
        if (!config.siteId().isBlank()) {
            JsonNode page = files.drives(config.siteId());
            page.path("value").forEach(node -> drives.add(mapper.toDrive(node, config.siteId())));
            return drives;
        }
        JsonNode meDrive = files.meDrive();
        drives.add(mapper.toDrive(meDrive));
        return drives;
    }

    private void crawlFolder(String driveId, String folderPath, ChangeSource source,
            String cursor, Map<String, Long> counts) throws IOException, InterruptedException {
        for (JsonNode node : files.childrenAll(driveId, folderPath)) {
            DriveItem item = mapper.toDriveItem(node, driveId);
            item = mapper.withListColumns(item,
                    files.listItemFieldsOrEmpty(driveId, item.getId()));
            if (includeContent && !item.getFolder() && item.getSize() > 0
                    && item.getSize() <= attachmentMaxBytes) {
                byte[] bytes = files.download(driveId, item.getId());
                if (bytes.length <= attachmentMaxBytes) {
                    item = item.toBuilder().setContent(ByteString.copyFrom(bytes)).build();
                }
            }
            emit(entity(item.getId()).setDriveItem(item).build(), source, cursor);
            counts.merge(item.getFolder() ? "folder" : "file", 1L, Long::sum);
            if (item.getFolder()) {
                String childPath = joinPath(folderPath, item.getName());
                crawlFolder(driveId, childPath, source, cursor, counts);
            }
        }
    }

    private void emit(MicrosoftEntity entity, ChangeSource source, String cursor) {
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId(UUID.randomUUID().toString())
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(entity)
                .setCursor(cursor)
                .setSource(source)
                .setOccurredAt(timestamp(Instant.now()))
                .build();
        MicrosoftValidator.create().requireValid(change);
        sink.emit(change);
    }

    private static MicrosoftEntity.Builder entity(String entityId) {
        return MicrosoftEntity.newBuilder()
                .setEntityId(entityId)
                .setIngestedAt(timestamp(Instant.now()));
    }

    private static String joinPath(String parent, String name) {
        if (parent == null || parent.isBlank() || parent.equals("/")) {
            return "/" + name;
        }
        return parent.endsWith("/") ? parent + name : parent + "/" + name;
    }

    private static Timestamp timestamp(Instant instant) {
        return Timestamp.newBuilder()
                .setSeconds(instant.getEpochSecond())
                .setNanos(instant.getNano())
                .build();
    }
}
