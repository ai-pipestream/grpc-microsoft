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
    }
}
