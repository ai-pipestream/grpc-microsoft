package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.DeltaPreferences;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.microsoft.v1.SyncCursor;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftCrawlerTest {

    private FakeGraphServer fake;
    private MicrosoftConnectorConfig config;
    private GraphFiles files;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        config = MicrosoftConnectorConfig.builder()
                .tenantId("t")
                .clientId("c")
                .clientSecret("s")
                .graphBaseUrl(fake.baseUrl())
                .build();
        files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
    }

    @AfterEach
    void stop() {
        fake.close();
    }

    private String link(String token) {
        return fake.baseUrl() + "/drives/drive-1/root/delta?token=" + token;
    }

    @Test
    void firstPassIsAFullDeltaEnumerationAcrossPages() throws Exception {
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.pageJson(link("page2"),
                MicrosoftFixtures.rootJson("drive-1"),
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1"),
                MicrosoftFixtures.folderJson("folder-1", "Shared", "drive-1")));
        fake.stub("/drives/drive-1/root/delta?token=page2", MicrosoftFixtures.deltaJson(link("d1"),
                MicrosoftFixtures.fileJson("file-2", "nested.txt", "drive-1")));
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields", """
                {"id":"7","contentType":{"id":"0x0101","name":"Document"},
                 "fields":{"Title":"Notes","Count":2,"@odata.etag":"e"}}
                """);

        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        MicrosoftCrawler.Result result = new MicrosoftCrawler(config, files, sink).crawl();

        FakeGraphServer.RecordedRequest first = fake.requestsTo("/drives/drive-1/root/delta").get(0);
        assertThat(first.query()).isEqualTo("$select=" + GraphPaths.DRIVE_ITEM_SELECT);
        assertThat(first.prefer()).isNull();
        assertThat(fake.requestsTo("/drives/drive-1/root/delta")).extracting(r -> r.query())
                .containsExactly("$select=" + GraphPaths.DRIVE_ITEM_SELECT, "token=page2");
        assertThat(fake.requestsTo("/drives/drive-1/root/children")).isEmpty();

        assertThat(sink.changes()).extracting(c -> c.getEntity().getEntityId())
                .containsExactly("drive-1", "root", "file-1", "folder-1", "file-2");
        assertThat(sink.changes()).allMatch(c -> c.getSource() == ChangeSource.CHANGE_SOURCE_CRAWL)
                .allMatch(c -> c.getOperation() == ChangeOperation.CHANGE_OPERATION_UPSERT)
                .allMatch(c -> c.getCursor().equals(result.runId()));
        MicrosoftSnapshot snapshot = sink.snapshots().get(0);
        assertThat(snapshot.getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_CRAWL);
        assertThat(snapshot.getCursor()).isEqualTo(link("d1"));
        assertThat(snapshot.getEntityCountsMap()).containsEntry("drive", 1L)
                .containsEntry("file", 2L).containsEntry("folder", 1L).containsEntry("root", 1L);
        assertThat(result.fullSweep()).isTrue();
        assertThat(result.cursor().getDriveDeltaLinksMap()).containsEntry("drive-1", link("d1"));

        DriveItem notes = item(sink, "file-1");
        assertThat(notes.getListColumnsList()).extracting(c -> c.getName())
                .containsExactly("Title", "Count");
        assertThat(notes.getListItemId()).isEqualTo("7");
        assertThat(notes.getContentType().getName()).isEqualTo("Document");
        assertThat(item(sink, "file-2").getListColumnsList()).isEmpty();
        assertThat(item(sink, "folder-1").getChildCount()).isEqualTo(1);
    }

    @Test
    void incrementalPassEmitsOnlyChangesAndDeletes() throws Exception {
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta?token=d1", MicrosoftFixtures.deltaJson(link("d2"),
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1"),
                MicrosoftFixtures.deletedJson("file-2", "drive-1")));
        SyncCursor since = SyncCursor.newBuilder().putDriveDeltaLinks("drive-1", link("d1")).build();
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();

        MicrosoftCrawler.Result result = new MicrosoftCrawler(config, files, sink,
                options(since, DeltaPreferences.newBuilder().setHierarchicalSharing(true)
                        .setShowRemovedAsDeleted(true).build())).crawl();

        FakeGraphServer.RecordedRequest delta = fake.only("/drives/drive-1/root/delta");
        assertThat(delta.query()).isEqualTo("token=d1");
        assertThat(delta.prefer()).isEqualTo("hierarchicalsharing, deltashowremovedasdeleted");
        MicrosoftChange deleted = sink.changes().stream()
                .filter(c -> c.getEntity().getEntityId().equals("file-2")).findFirst().orElseThrow();
        assertThat(deleted.getOperation()).isEqualTo(ChangeOperation.CHANGE_OPERATION_DELETE);
        assertThat(deleted.getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_DELTA);
        assertThat(deleted.getCursor()).isEqualTo(link("d1"));
        assertThat(deleted.getEntity().getDriveItem().hasDeleted()).isTrue();
        assertThat(fake.requestsTo("/drives/drive-1/items/file-2/listItem")).isEmpty();
        assertThat(sink.snapshots().get(0).getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_DELTA);
        assertThat(sink.snapshots().get(0).getEntityCountsMap()).containsEntry("deleted", 1L);
        assertThat(result.fullSweep()).isFalse();
        assertThat(result.cursor().getDriveDeltaLinksMap()).containsEntry("drive-1", link("d2"));
    }

    @Test
    void goneDeltaLinkRestartsTheDriveAsAFullSweep() throws Exception {
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/delta?token=stale", FakeGraphServer.error(410,
                "resyncChangesApplyDifferences", Map.of("Location", link("fresh"))));
        fake.stub("/drives/drive-1/root/delta?token=fresh", MicrosoftFixtures.deltaJson(link("d3"),
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")));
        SyncCursor since = SyncCursor.newBuilder().putDriveDeltaLinks("drive-1", link("stale")).build();
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();

        MicrosoftCrawler.Result result = new MicrosoftCrawler(config, files, sink,
                options(since, DeltaPreferences.getDefaultInstance())).crawl();

        assertThat(fake.requestsTo("/drives/drive-1/root/delta")).extracting(r -> r.query())
                .containsExactly("token=stale", "token=fresh");
        MicrosoftChange file = sink.changes().stream()
                .filter(c -> c.getEntity().getEntityId().equals("file-1")).findFirst().orElseThrow();
        assertThat(file.getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_CRAWL);
        assertThat(file.getCursor()).isEqualTo(result.runId());
        assertThat(sink.snapshots().get(0).getSource()).isEqualTo(ChangeSource.CHANGE_SOURCE_CRAWL);
        assertThat(result.fullSweep()).isTrue();
        assertThat(result.cursor().getDriveDeltaLinksMap()).containsEntry("drive-1", link("d3"));
    }

    @Test
    void folderPathFiltersTheWholeDriveFeedByAncestry() throws Exception {
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root:/Shared:", "{\"id\":\"folder-1\"}");
        fake.stub("/drives/drive-1/root/delta", MicrosoftFixtures.deltaJson(link("d1"),
                MicrosoftFixtures.rootJson("drive-1"),
                MicrosoftFixtures.fileJson("file-1", "top.txt", "drive-1"),
                MicrosoftFixtures.folderJson("folder-1", "Shared", "drive-1"),
                """
                {"id":"file-2","name":"nested.txt","size":3,"file":{"mimeType":"text/plain"},
                 "parentReference":{"id":"folder-1","driveId":"drive-1"}}
                """,
                """
                {"id":"file-3","name":"deep.txt","size":3,"file":{"mimeType":"text/plain"},
                 "parentReference":{"id":"sub-1","driveId":"drive-1"}}
                """));
        fake.stub("/drives/drive-1/items/sub-1",
                "{\"id\":\"sub-1\",\"parentReference\":{\"id\":\"folder-1\"}}");
        MicrosoftConnectorConfig scoped = MicrosoftConnectorConfig.builder()
                .tenantId("t").clientId("c").clientSecret("s").graphBaseUrl(fake.baseUrl())
                .folderPath("/Shared").build();
        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();

        new MicrosoftCrawler(scoped, files, sink).crawl();

        assertThat(sink.changes()).extracting(c -> c.getEntity().getEntityId())
                .containsExactly("drive-1", "file-2", "file-3");
        assertThat(fake.only("/drives/drive-1/items/sub-1").query())
                .isEqualTo("$select=id,parentReference");
    }

    private MicrosoftCrawler.Options options(SyncCursor since, DeltaPreferences prefs) {
        return new MicrosoftCrawler.Options(List.of(), List.of(), "", false, 0, true, false, false,
                false, false, false, false, prefs, false, since);
    }

    private static DriveItem item(InMemoryMicrosoftChangeSink sink, String id) {
        return sink.changes().stream()
                .filter(c -> id.equals(c.getEntity().getEntityId()))
                .findFirst().orElseThrow()
                .getEntity().getDriveItem();
    }
}
