package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.FileHashes;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.KnowledgeBundle;
import ai.pipestream.okf.OkfConformance;
import ai.pipestream.okf.OkfOutput;
import ai.pipestream.okf.warc.WarcArchive;
import ai.pipestream.okf.warc.WarcParse;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.GZIPInputStream;
import java.util.zip.ZipFile;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * End-to-end OKF + WARC shape a Microsoft Graph crawl should present: download
 * URL as WARC-Target-URI, SharePoint columns in markdown (not JSON), zip as a
 * sibling never nested in WARC.
 */
class OkfMicrosoftWarcPresentationTest {

    private static Timestamp ts() {
        return Timestamp.newBuilder().setSeconds(1_700_000_000).build();
    }

    @Test
    void graphBundleHasSiblingWarcNotZipInsideWarc(@TempDir Path dir) throws Exception {
        byte[] notes = "hello notes".getBytes(StandardCharsets.UTF_8);
        List<CatalogEntry> entries = new ArrayList<>();
        entries.add(MicrosoftCatalog.from(change(MicrosoftEntity.newBuilder()
                .setEntityId("site-1")
                .setSite(Site.newBuilder().setId("site-1").setName("ENG")
                        .setDisplayName("Engineering")
                        .setWebUrl("https://contoso.sharepoint.com/sites/ENG")))).orElseThrow());
        entries.add(MicrosoftCatalog.from(change(MicrosoftEntity.newBuilder()
                .setEntityId("drive-1")
                .setDrive(Drive.newBuilder().setId("drive-1").setName("Docs")
                        .setDriveType("documentLibrary")
                        .setWebUrl("https://contoso.sharepoint.com/sites/ENG/Docs")))).orElseThrow());
        entries.add(MicrosoftCatalog.from(change(MicrosoftEntity.newBuilder()
                .setEntityId("folder-1")
                .setDriveItem(DriveItem.newBuilder()
                        .setId("folder-1").setName("Shared").setDriveId("drive-1")
                        .setFolder(true).setChildCount(1)
                        .setWebUrl("https://contoso.sharepoint.com/sites/ENG/Shared")))).orElseThrow());
        entries.add(MicrosoftCatalog.from(change(MicrosoftEntity.newBuilder()
                .setEntityId("file-1")
                .setDriveItem(DriveItem.newBuilder()
                        .setId("file-1").setName("notes.txt").setDriveId("drive-1")
                        .setMimeType("text/plain")
                        .setWebUrl("https://contoso.sharepoint.com/sites/ENG/notes.txt")
                        .setDownloadUrl("https://contoso.sharepoint.com/download/notes.txt")
                        .setEtag("etag-1")
                        .setHashes(FileHashes.newBuilder().setSha256("abc"))
                        .addListColumns(ListColumn.newBuilder().setName("Title")
                                .setStringValue("Notes"))
                        .setContent(ByteString.copyFrom(notes))))).orElseThrow());

        KnowledgeBundle bundle = KnowledgeBundle.assemble(
                "Engineering library",
                "OKF v0.2 capture of a MicrosoftService.Sync run.",
                MicrosoftCatalog.ACTOR,
                Instant.parse("2026-08-23T12:00:00Z"),
                "grpc-microsoft/okf-producer",
                entries);
        Path tree = dir.resolve("eng");
        Path zip = dir.resolve("eng.zip");
        Path warcGz = dir.resolve("eng.warc.gz");
        bundle.write(new OkfOutput(tree, zip, warcGz));

        assertThat(OkfConformance.check(bundle.okf())).isEmpty();
        assertThat(Files.readString(tree.resolve("index.md"))).contains("okf_version: \"0.2\"");
        assertThat(tree.resolve("sites/site-1.md")).exists();
        assertThat(tree.resolve("drives/drive-1.md")).exists();
        assertThat(tree.resolve("items/drive-1/folder-1.md")).exists();
        assertThat(tree.resolve("items/drive-1/file-1.md")).exists();
        String fileMd = Files.readString(tree.resolve("items/drive-1/file-1.md"));
        assertThat(fileMd).contains("| Title | Notes |").contains("SHA-256: `abc`");
        assertThat(fileMd).doesNotContain("\"Title\": \"Notes\"");

        String collection = Files.readString(tree.resolve("collection.html"));
        assertThat(collection).contains("href=\"https://contoso.sharepoint.com/download/notes.txt\"");
        assertThat(collection).contains("href=\"https://contoso.sharepoint.com/sites/ENG\"");
        assertThat(collection).doesNotContain("href=\"items/drive-1/file-1.md\"");

        try (ZipFile zipFile = new ZipFile(zip.toFile())) {
            assertThat(zipFile.getEntry("items/drive-1/file-1.md")).isNotNull();
            assertThat(zipFile.getEntry("collection.html")).isNotNull();
        }

        byte[] warcBytes;
        try (GZIPInputStream in = new GZIPInputStream(Files.newInputStream(warcGz))) {
            warcBytes = in.readAllBytes();
        }
        List<WarcParse.Record> records = WarcParse.parse(warcBytes);
        assertThat(records.get(0).type()).isEqualTo("warcinfo");
        assertThat(records.get(0).headers()).doesNotContainKey("WARC-Target-URI");

        List<WarcParse.Record> resources = records.stream()
                .filter(r -> "resource".equals(r.type())).toList();
        List<WarcParse.Record> conversions = records.stream()
                .filter(r -> "conversion".equals(r.type())).toList();
        assertThat(conversions).hasSize(entries.size());
        assertThat(resources).hasSize(entries.size() + 1);

        assertThat(resources).extracting(WarcParse.Record::targetUri)
                .containsAll(entries.stream().map(CatalogEntry::targetUri).toList())
                .contains(WarcArchive.COLLECTION_URI);

        WarcParse.Record fileResource = records.stream()
                .filter(r -> "https://contoso.sharepoint.com/download/notes.txt".equals(r.targetUri()))
                .findFirst().orElseThrow();
        assertThat(fileResource.block()).isEqualTo(notes);

        String warcText = new String(warcBytes, StandardCharsets.ISO_8859_1);
        assertThat(warcText).doesNotContain("PK\u0003\u0004");
        assertThat(Files.size(zip)).isPositive();
        assertThat(warcBytes.length).isNotEqualTo(Files.size(zip));
    }

    @Test
    void changeWithoutEntityIsSkipped() {
        assertThat(MicrosoftCatalog.from(MicrosoftChange.newBuilder().setChangeId("x").build()))
                .isEmpty();
    }

    private static MicrosoftChange change(MicrosoftEntity.Builder entity) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setEntity(entity.setIngestedAt(ts()))
                .build();
    }
}
