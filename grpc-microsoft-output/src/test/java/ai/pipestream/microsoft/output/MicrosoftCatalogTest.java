package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.FileHashes;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.OkfConcept;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftCatalogTest {

    private static Timestamp ts() {
        return Timestamp.newBuilder().setSeconds(1).build();
    }

    @Test
    void fileUsesDownloadUrlAndRendersColumns() {
        DriveItem item = DriveItem.newBuilder()
                .setId("file-1")
                .setName("notes.txt")
                .setDriveId("drive-1")
                .setWebUrl("https://contoso.sharepoint.com/notes.txt")
                .setDownloadUrl("https://contoso.sharepoint.com/download/notes.txt")
                .setMimeType("text/plain")
                .setDescription("a note")
                .setEtag("etag-1")
                .setHashes(FileHashes.newBuilder().setSha256("abc"))
                .addListColumns(ListColumn.newBuilder().setName("Title").setStringValue("Notes"))
                .addListColumns(ListColumn.newBuilder().setName("Count").setIntValue(3))
                .build();
        CatalogEntry entry = MicrosoftCatalog.from(change(item)).orElseThrow();
        assertThat(entry.path()).isEqualTo("items/drive-1/file-1.md");
        assertThat(entry.targetUri()).isEqualTo("https://contoso.sharepoint.com/download/notes.txt");
        assertThat(entry.concept().body()).contains("| Title | Notes |").contains("| Count | 3 |")
                .contains("eTag: `etag-1`").contains("SHA-256: `abc`");
    }

    @Test
    void fileBodyRendersEveryHashAndEscapesPipes() {
        CatalogEntry entry = MicrosoftCatalog.from(change(DriveItem.newBuilder()
                .setId("file-1").setName("a|b.txt").setDriveId("drive-1")
                .setHashes(FileHashes.newBuilder()
                        .setSha1("s1").setSha256("s2").setQuickXor("qx").setCrc32("c32"))
                .addListColumns(ListColumn.newBuilder().setName("Choice").setStringValue("a|b"))
                .addListColumns(ListColumn.newBuilder().setName("Flag").setBoolValue(true))
                .addListColumns(ListColumn.newBuilder().setName("Ratio").setDoubleValue(1.5d))
                .build())).orElseThrow();
        assertThat(entry.concept().body())
                .contains("SHA-1: `s1`")
                .contains("SHA-256: `s2`")
                .contains("QuickXOR: `qx`")
                .contains("CRC32: `c32`")
                .contains("| Choice | a\\|b |")
                .contains("| Flag | true |")
                .contains("| Ratio | 1.5 |");
    }

    @Test
    void emptyChangeIsSkipped() {
        assertThat(MicrosoftCatalog.from(MicrosoftChange.newBuilder().setChangeId("x").build()))
                .isEmpty();
    }

    @Test
    void folderSiteDriveUserAndDelete() {
        CatalogEntry folder = MicrosoftCatalog.from(change(DriveItem.newBuilder()
                .setId("folder-1").setName("Shared").setDriveId("drive-1")
                .setFolder(true).setChildCount(2)
                .setWebUrl("https://contoso.sharepoint.com/Shared").build())).orElseThrow();
        assertThat(folder.kind()).isEqualTo("folder");
        assertThat(folder.concept().type()).isEqualTo("Folder");
        assertThat(folder.concept().body()).contains("Child count: 2");

        CatalogEntry site = MicrosoftCatalog.from(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(MicrosoftEntity.newBuilder().setEntityId("site-1").setIngestedAt(ts())
                        .setSite(Site.newBuilder().setId("site-1").setName("ENG")
                                .setDisplayName("Engineering")
                                .setWebUrl("https://contoso.sharepoint.com/sites/ENG")))
                .build()).orElseThrow();
        assertThat(site.path()).isEqualTo("sites/site-1.md");

        CatalogEntry drive = MicrosoftCatalog.from(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(MicrosoftEntity.newBuilder().setEntityId("drive-1").setIngestedAt(ts())
                        .setDrive(Drive.newBuilder().setId("drive-1").setName("Docs")
                                .setDriveType("documentLibrary")))
                .build()).orElseThrow();
        assertThat(drive.targetUri()).startsWith("urn:okf:0.2:microsoft:drive:");

        CatalogEntry user = MicrosoftCatalog.from(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(MicrosoftEntity.newBuilder().setEntityId("u1").setIngestedAt(ts())
                        .setUser(GraphUser.newBuilder().setId("u1").setDisplayName("Ada")
                                .setMail("ada@contoso.com")))
                .build()).orElseThrow();
        assertThat(user.path()).isEqualTo("users/u1.md");
        assertThat(user.concept().extra()).containsEntry("mail", "ada@contoso.com");

        CatalogEntry deleted = MicrosoftCatalog.from(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setEntity(MicrosoftEntity.newBuilder().setEntityId("file-1").setIngestedAt(ts())
                        .setDriveItem(DriveItem.newBuilder().setId("file-1").setName("gone")
                                .setDriveId("d")))
                .build()).orElseThrow();
        assertThat(deleted.concept().status()).contains(OkfConcept.Status.DEPRECATED);
    }

    @Test
    void inlineContentBecomesResourceBody() {
        CatalogEntry entry = MicrosoftCatalog.from(change(DriveItem.newBuilder()
                .setId("f").setName("a.txt").setDriveId("d")
                .setMimeType("text/plain")
                .setWebUrl("https://contoso/a.txt")
                .setContent(ByteString.copyFromUtf8("hello"))
                .build())).orElseThrow();
        assertThat(new String(entry.resourceBody())).isEqualTo("hello");
        assertThat(entry.mediaType()).isEqualTo("text/plain");
        assertThat(entry.concept().extra()).containsKey("sha256");
    }

    private static MicrosoftChange change(DriveItem item) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(item.getId())
                        .setIngestedAt(ts())
                        .setDriveItem(item))
                .build();
    }
}
