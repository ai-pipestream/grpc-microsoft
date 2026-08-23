package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.DriveItem;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

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

    @Test
    void crawlsDriveAndNestedFolder() throws Exception {
        fake.stub("/me/drive", MicrosoftFixtures.driveJson("drive-1", "Docs"));
        fake.stub("/drives/drive-1/root/children", MicrosoftFixtures.childrenJson(
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1"),
                MicrosoftFixtures.folderJson("folder-1", "Shared", "drive-1")));
        fake.stub("/drives/drive-1/root:/Shared:/children", MicrosoftFixtures.childrenJson(
                MicrosoftFixtures.fileJson("file-2", "nested.txt", "drive-1")));
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields", """
                {"fields":{"Title":"Notes","Count":2,"@odata.etag":"e"}}
                """);

        InMemoryMicrosoftChangeSink sink = new InMemoryMicrosoftChangeSink();
        new MicrosoftCrawler(config, files, sink).crawl();

        assertThat(sink.changes()).extracting(c -> c.getEntity().getEntityId())
                .contains("drive-1", "file-1", "folder-1", "file-2");
        assertThat(sink.snapshots()).hasSize(1);
        assertThat(sink.snapshots().get(0).getEntityCountsMap())
                .containsEntry("drive", 1L)
                .containsEntry("file", 2L)
                .containsEntry("folder", 1L);
        DriveItem notes = sink.changes().stream()
                .filter(c -> "file-1".equals(c.getEntity().getEntityId()))
                .findFirst().orElseThrow()
                .getEntity().getDriveItem();
        assertThat(notes.getListColumnsList()).extracting(c -> c.getName())
                .containsExactly("Title", "Count");
        assertThat(notes.getListColumns(1).getIntValue()).isEqualTo(2);
    }
}
