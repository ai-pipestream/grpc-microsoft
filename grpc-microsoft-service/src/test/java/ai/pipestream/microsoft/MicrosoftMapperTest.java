package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.Site;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftMapperTest {

    private static final ObjectMapper JSON = new ObjectMapper();
    private final MicrosoftMapper mapper = new MicrosoftMapper();

    @Test
    void mapsUserSiteDriveAndFile() throws Exception {
        GraphUser user = mapper.toUser(JSON.readTree(MicrosoftFixtures.meJson()));
        assertThat(user.getId()).isEqualTo("user-1");
        assertThat(user.getMail()).isEqualTo("bot@contoso.com");

        Site site = mapper.toSite(JSON.readTree(MicrosoftFixtures.siteJson("site-1", "ENG")));
        assertThat(site.getDisplayName()).isEqualTo("ENG");

        Drive drive = mapper.toDrive(JSON.readTree(MicrosoftFixtures.driveJson("drive-1", "Docs")),
                "site-1");
        assertThat(drive.getDriveType()).isEqualTo("business");
        assertThat(drive.getSiteId()).isEqualTo("site-1");

        DriveItem file = mapper.toDriveItem(
                JSON.readTree(MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1")),
                "drive-1");
        assertThat(file.getMimeType()).isEqualTo("text/plain");
        assertThat(file.getFolder()).isFalse();
        assertThat(file.getParentId()).isEqualTo("root");
        assertThat(file.getCreatedBy().getUser().getDisplayName()).isEqualTo("Bot");
        assertThat(file.getCreatedAt().getSeconds()).isGreaterThan(0);
    }

    @Test
    void mapsFolder() throws Exception {
        DriveItem folder = mapper.toDriveItem(
                JSON.readTree(MicrosoftFixtures.folderJson("folder-1", "Shared", "drive-1")),
                "drive-1");
        assertThat(folder.getFolder()).isTrue();
        assertThat(folder.getMimeType()).isEmpty();
        assertThat(folder.getChildCount()).isEqualTo(1);
    }

    @Test
    void flattensSharePointColumnsSkippingOdata() throws Exception {
        DriveItem item = mapper.withListColumns(DriveItem.newBuilder().setId("file-1").build(),
                JSON.readTree("""
                        {
                          "Title": "Notes",
                          "Count": 3,
                          "Flag": true,
                          "When": "2024-03-02T00:00:00Z",
                          "Tags": ["a", "b"],
                          "Lookup": {"Id": 9, "Label": "X"},
                          "@odata.etag": "skip-me",
                          "@odata.context": "also-skip"
                        }
                        """));
        assertThat(item.getListColumnsList()).extracting(c -> c.getName())
                .containsExactly("Title", "Count", "Flag", "When", "Tags", "Lookup.Id",
                        "Lookup.Label");
        assertThat(item.getListColumnsList()).noneMatch(c -> c.getName().startsWith("@odata"));
        assertThat(item.getListColumns(0).getStringValue()).isEqualTo("Notes");
        assertThat(item.getListColumns(1).getIntValue()).isEqualTo(3);
        assertThat(item.getListColumns(2).getBoolValue()).isTrue();
        assertThat(item.getListColumns(3).hasTimestampValue()).isTrue();
        assertThat(item.getListColumns(4).getStringValue()).isEqualTo("a, b");
        assertThat(item.getListColumns(5).getIntValue()).isEqualTo(9);
    }

    @Test
    void mapsHashesEtagDescriptionAndDownloadUrl() throws Exception {
        DriveItem file = mapper.toDriveItem(
                JSON.readTree(MicrosoftFixtures.fileJsonWithHashes("file-1", "notes.txt", "drive-1")),
                "drive-1");
        assertThat(file.getDescription()).isEqualTo("a note");
        assertThat(file.getEtag()).isEqualTo("etag-1");
        assertThat(file.getDownloadUrl()).isEqualTo("https://contoso.sharepoint.com/download/notes.txt");
        assertThat(file.getHashes().getSha1()).isEqualTo("aaa");
        assertThat(file.getHashes().getSha256()).isEqualTo("bbb");
        assertThat(file.getHashes().getQuickXor()).isEqualTo("ccc");
        assertThat(file.getHashes().getCrc32()).isEqualTo("ddd");
    }

    @Test
    void listColumnsSkipNullEmptyArrayAndSecondNesting() throws Exception {
        assertThat(mapper.toListColumns(null)).isEmpty();
        assertThat(mapper.toListColumns(JSON.readTree("[]"))).isEmpty();
        assertThat(mapper.toListColumns(JSON.readTree("""
                {
                  "Title": null,
                  "Empty": [],
                  "Ratio": 1.5,
                  "Day": "2024-03-02",
                  "Lookup": {"Id": 9, "Nested": {"x": 1}},
                  "@skip": "annotation"
                }
                """))).satisfies(columns -> {
            assertThat(columns).extracting(c -> c.getName())
                    .containsExactly("Ratio", "Day", "Lookup.Id");
            assertThat(columns.get(0).getDoubleValue()).isEqualTo(1.5d);
            assertThat(columns.get(1).getStringValue()).isEqualTo("2024-03-02");
            assertThat(columns.get(2).getIntValue()).isEqualTo(9);
        });
    }

    @Test
    void parsesRfc3339Offsets() {
        assertThat(MicrosoftMapper.timestamp("2024-03-01T12:00:00Z").getSeconds())
                .isEqualTo(MicrosoftMapper.timestamp("2024-03-01T14:00:00+02:00").getSeconds());
        assertThat(MicrosoftMapper.timestamp("")).isNull();
        assertThat(MicrosoftMapper.timestamp("not-a-date")).isNull();
    }
}
