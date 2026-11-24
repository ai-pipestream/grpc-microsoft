package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.KnowledgeBundle;
import ai.pipestream.okf.OkfConformance;
import ai.pipestream.okf.warc.WarcArchive;
import ai.pipestream.okf.warc.WarcParse;
import ai.pipestream.okf.warc.WarcWriter;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.time.Instant;
import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The archive holds a record for a drive item's URL only when bytes read
 * from Graph stand behind it. A Sync without {@code include_content}, a file
 * over the inline limit, and every folder arrive as metadata: such an item
 * keeps its OKF markdown and its link, and gets no resource record and no
 * stand-in text. The item is named by its stable {@code web_url}, never by
 * the short-lived pre-authenticated {@code download_url}.
 */
class NotFetchedFileTest {

    private static final String WEB =
            "https://contoso.sharepoint.com/sites/ENG/Shared%20Documents/Quarterly%20report.docx";
    private static final String DOWNLOAD =
            "https://contoso.sharepoint.com/sites/ENG/_layouts/15/download.aspx"
                    + "?UniqueId=4f1c&tempauth=v1.eyJzaXRlaWQiOiJ4In0";
    private static final String DOCX =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    @Test
    void aFileWhoseBytesWereNotFetchedHasNoResourceRecord() {
        CatalogEntry entry = entry(file(), ChangeOperation.CHANGE_OPERATION_UPSERT);

        List<WarcParse.Record> records = warc(entry);

        // Nothing in the archive passes for the document: no record under
        // either of its URLs, and not its file name stored as its bytes.
        assertThat(records).extracting(WarcParse.Record::targetUri).doesNotContain(WEB, DOWNLOAD);
        assertThat(records).filteredOn(r -> "resource".equals(r.type()))
                .extracting(r -> new String(r.block(), StandardCharsets.UTF_8))
                .doesNotContain("Quarterly report.docx");
        // The file is still described and linked, and says it was not fetched.
        WarcParse.Record markdown = conversion(records, "items/drive-1/file-1.md");
        assertThat(new String(markdown.block(), StandardCharsets.UTF_8))
                .contains("type: File", "capture: not_fetched", "* Size: 48213", "* MIME: " + DOCX);
        assertThat(markdown.refersTo()).isNull();
        assertThat(collection(records)).contains("href=\"" + WEB + "\"");
    }

    @Test
    void thePreAuthenticatedDownloadUrlIsNeverWritten() {
        List<WarcParse.Record> metadataOnly =
                warc(entry(file(), ChangeOperation.CHANGE_OPERATION_UPSERT));
        List<WarcParse.Record> withBytes = warc(entry(
                file().setContent(ByteString.copyFromUtf8("PK-docx-bytes")),
                ChangeOperation.CHANGE_OPERATION_UPSERT));

        for (List<WarcParse.Record> records : List.of(metadataOnly, withBytes)) {
            assertThat(records).extracting(WarcParse.Record::targetUri).doesNotContain(DOWNLOAD);
            assertThat(records).extracting(r -> new String(r.block(), StandardCharsets.UTF_8))
                    .noneMatch(block -> block.contains("tempauth"));
        }
    }

    @Test
    void aFileWhoseBytesWereFetchedIsStoredUnderItsWebUrl() {
        byte[] bytes = "PK-docx-bytes".getBytes(StandardCharsets.UTF_8);
        CatalogEntry entry = entry(file().setContent(ByteString.copyFrom(bytes)),
                ChangeOperation.CHANGE_OPERATION_UPSERT);

        List<WarcParse.Record> records = warc(entry);

        WarcParse.Record resource = resource(records, WEB);
        assertThat(resource.block()).isEqualTo(bytes);
        assertThat(resource.headers().get("Content-Type")).isEqualTo(DOCX);
        WarcParse.Record markdown = conversion(records, "items/drive-1/file-1.md");
        assertThat(markdown.refersTo()).isEqualTo(resource.recordId());
        assertThat(new String(markdown.block(), StandardCharsets.UTF_8))
                .doesNotContain("capture:");
    }

    @Test
    void aFolderHasNoResourceRecord() {
        String url = "https://contoso.sharepoint.com/sites/ENG/Shared%20Documents/Reports";
        CatalogEntry entry = entry(DriveItem.newBuilder()
                .setId("folder-1").setName("Reports").setDriveId("drive-1")
                .setFolder(true).setChildCount(3).setWebUrl(url),
                ChangeOperation.CHANGE_OPERATION_UPSERT);

        List<WarcParse.Record> records = warc(entry);

        assertThat(records).extracting(WarcParse.Record::targetUri).doesNotContain(url);
        assertThat(new String(conversion(records, "items/drive-1/folder-1.md").block(),
                StandardCharsets.UTF_8))
                .contains("type: Folder", "capture: not_fetched", "* Child count: 3");
        assertThat(collection(records)).contains("href=\"" + url + "\"");
    }

    @Test
    void aFileWithOnlyADownloadUrlIsNamedByItsIds() {
        CatalogEntry entry = entry(file().clearWebUrl(), ChangeOperation.CHANGE_OPERATION_UPSERT);

        assertThat(entry.targetUri()).startsWith("urn:okf:0.2:microsoft:item:")
                .contains("drive-1").contains("file-1");
    }

    private static DriveItem.Builder file() {
        return DriveItem.newBuilder()
                .setId("file-1").setName("Quarterly report.docx").setDriveId("drive-1")
                .setMimeType(DOCX)
                .setSize(48_213)
                .setWebUrl(WEB)
                .setDownloadUrl(DOWNLOAD);
    }

    private static CatalogEntry entry(DriveItem.Builder item, ChangeOperation operation) {
        return MicrosoftCatalog.from(MicrosoftChange.newBuilder()
                .setChangeId("c")
                .setOperation(operation)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(item.getId())
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1_700_000_000))
                        .setDriveItem(item))
                .build()).orElseThrow();
    }

    private static List<WarcParse.Record> warc(CatalogEntry... entries) {
        KnowledgeBundle bundle = KnowledgeBundle.assemble(
                "Engineering library",
                "OKF v0.2 capture of a MicrosoftService.Sync run.",
                MicrosoftCatalog.ACTOR,
                Instant.parse("2026-08-23T12:00:00Z"),
                "grpc-microsoft/okf-producer",
                List.of(entries));
        assertThat(OkfConformance.check(bundle.okf())).isEmpty();
        return WarcParse.parse(WarcWriter.toBytes(bundle.warcRecords()));
    }

    private static WarcParse.Record resource(List<WarcParse.Record> records, String uri) {
        return records.stream()
                .filter(r -> "resource".equals(r.type()) && uri.equals(r.targetUri()))
                .findFirst().orElseThrow();
    }

    private static WarcParse.Record conversion(List<WarcParse.Record> records, String path) {
        return records.stream()
                .filter(r -> "conversion".equals(r.type())
                        && ("urn:okf:0.2:" + path).equals(r.targetUri()))
                .findFirst().orElseThrow();
    }

    private static String collection(List<WarcParse.Record> records) {
        return new String(resource(records, WarcArchive.COLLECTION_URI).block(),
                StandardCharsets.UTF_8);
    }
}
