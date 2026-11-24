package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.Identity;
import ai.pipestream.microsoft.v1.IdentitySet;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import microsoft.graph.connectors.contracts.grpc.Content;
import microsoft.graph.connectors.contracts.grpc.CrawlItem;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlItem;
import org.junit.jupiter.api.Test;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class CrawlItemMapperTest {

    @Test
    void mapsFileToContentItem() {
        DriveItem item = DriveItem.newBuilder()
                .setId("file-1")
                .setName("notes.txt")
                .setDriveId("drive-1")
                .setWebUrl("https://contoso/notes.txt")
                .setMimeType("text/plain")
                .setSize(5)
                .setContent(ByteString.copyFromUtf8("hello"))
                .setCreatedBy(IdentitySet.newBuilder()
                        .setUser(Identity.newBuilder().setId("u1").setDisplayName("Ada")))
                .build();

        CrawlItemMapper.Mapped<CrawlItem> mapped = CrawlItemMapper.toCrawlItem(item, true);
        CrawlItem crawl = mapped.item();

        assertThat(mapped.omittedContent()).isEmpty();
        assertThat(crawl.getItemType()).isEqualTo(CrawlItem.ItemType.ContentItem);
        assertThat(crawl.getItemId()).isEqualTo("drive-1/file-1");
        assertThat(crawl.getContentItem().getPropertyValues().getValuesMap()
                .get(DataSourceSchemas.TITLE).getStringValue()).isEqualTo("notes.txt");
        assertThat(crawl.getContentItem().getPropertyValues().getValuesMap()
                .get(DataSourceSchemas.FILE_EXTENSION).getStringValue()).isEqualTo("txt");
        assertThat(crawl.getContentItem().getPropertyValues().getValuesMap()
                .get(DataSourceSchemas.CREATED_BY).getPrincipalValue().getExternalName())
                .isEqualTo("Ada");
        assertThat(crawl.getContentItem().getContent().getContentType())
                .isEqualTo(Content.ContentType.Text);
        assertThat(crawl.getContentItem().getContent().getContentValue()).isEqualTo("hello");
    }

    @Test
    void deleteBecomesIncrementalDeletedItemUnderTheIdItWasIndexedWith() {
        DriveItem indexed = DriveItem.newBuilder().setId("file-1").setDriveId("drive-1")
                .setName("notes.txt").build();
        MicrosoftChange change = delete(MicrosoftEntity.newBuilder()
                .setEntityId("file-1")
                .setDriveItem(DriveItem.newBuilder().setId("file-1").setDriveId("drive-1")
                        .setDeleted(ai.pipestream.microsoft.v1.Deleted.newBuilder()
                                .setState("deleted"))));

        IncrementalCrawlItem item =
                CrawlItemMapper.toIncrementalCrawlItem(change, true).orElseThrow().item();

        assertThat(item.getItemType()).isEqualTo(IncrementalCrawlItem.ItemType.DeletedItem);
        assertThat(item.getItemId()).isEqualTo("drive-1/file-1")
                .isEqualTo(CrawlItemMapper.toCrawlItem(indexed, true).item().getItemId());
        assertThat(CrawlItemMapper.toCrawlItem(change, true)).isEmpty();
    }

    @Test
    void aDeleteWithoutADriveIdUsesTheIdAnUpsertWithoutOneWouldHave() {
        DriveItem indexed = DriveItem.newBuilder().setId("file-1").build();
        MicrosoftChange change = delete(MicrosoftEntity.newBuilder()
                .setEntityId("file-1").setDriveItem(indexed));

        assertThat(CrawlItemMapper.toIncrementalCrawlItem(change, true).orElseThrow().item()
                .getItemId())
                .isEqualTo("file-1")
                .isEqualTo(CrawlItemMapper.toCrawlItem(indexed, true).item().getItemId());
    }

    @Test
    void aDeleteThatDoesNotCarryItsDriveItemIsRefusedNotSentUnderAnotherId() {
        MicrosoftChange bare = delete(MicrosoftEntity.newBuilder().setEntityId("gone"));
        MicrosoftChange noId = delete(MicrosoftEntity.newBuilder().setEntityId("gone")
                .setDriveItem(DriveItem.newBuilder().setDriveId("drive-1")));

        assertThatThrownBy(() -> CrawlItemMapper.toIncrementalCrawlItem(bare, true))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("c1").hasMessageContaining("gone");
        assertThatThrownBy(() -> CrawlItemMapper.toIncrementalCrawlItem(noId, true))
                .isInstanceOf(IllegalArgumentException.class)
                .hasMessageContaining("c1").hasMessageContaining("gone");
    }

    @Test
    void aDeleteOfAnEntityThatIsNeverIndexedIsSkippedLikeItsUpsert() {
        MicrosoftChange change = delete(MicrosoftEntity.newBuilder()
                .setEntityId("list-item-1")
                .setListItem(ai.pipestream.microsoft.v1.ListItem.newBuilder().setId("7")));

        assertThat(CrawlItemMapper.toIncrementalCrawlItem(change, true)).isEmpty();
    }

    private static MicrosoftChange delete(MicrosoftEntity.Builder entity) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setEntity(entity.setIngestedAt(Timestamp.newBuilder().setSeconds(1)))
                .build();
    }

    @Test
    void skipsNonDriveItemEntities() {
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId("c2")
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("drive-1")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1))
                        .setDrive(ai.pipestream.microsoft.v1.Drive.newBuilder().setId("drive-1")))
                .build();

        assertThat(CrawlItemMapper.toCrawlItem(change, true)).isEmpty();
    }

    @Test
    void textAtTheLimitIsSentAndOneByteMoreIsLeftOut() {
        byte[] atLimit = new byte[CrawlItemMapper.CONTENT_CAP_BYTES];
        java.util.Arrays.fill(atLimit, (byte) 'a');
        byte[] over = new byte[CrawlItemMapper.CONTENT_CAP_BYTES + 1];
        java.util.Arrays.fill(over, (byte) 'a');

        assertThat(map("text/plain", atLimit).item().getContentItem().getContent()
                .getContentValue()).hasSize(CrawlItemMapper.CONTENT_CAP_BYTES);
        CrawlItemMapper.Mapped<CrawlItem> mapped = map("text/plain", over);
        assertThat(mapped.item().getContentItem().hasContent()).isFalse();
        assertThat(mapped.omittedContent().orElseThrow().reason())
                .isEqualTo(OmittedContent.Reason.TOO_LARGE);
    }

    @Test
    void theLimitIsOnTheTextAsItIsSentNotOnTheFileBytes() {
        // Each of these characters is two bytes as UTF-16 and three as UTF-8:
        // the file is under the limit, the text GCA would receive is over it.
        String text = "\u4e2d".repeat(CrawlItemMapper.CONTENT_CAP_BYTES / 2 - 4);
        byte[] utf16 = ("\ufeff" + text).getBytes(StandardCharsets.UTF_16LE);
        assertThat(utf16.length).isLessThan(CrawlItemMapper.CONTENT_CAP_BYTES);

        CrawlItemMapper.Mapped<CrawlItem> mapped = map("text/plain", utf16);

        assertThat(mapped.item().getContentItem().hasContent()).isFalse();
        OmittedContent omitted = mapped.omittedContent().orElseThrow();
        assertThat(omitted.reason()).isEqualTo(OmittedContent.Reason.TOO_LARGE);
        assertThat(omitted.detail()).contains("UTF-8");
    }

    @Test
    void aUtf32ByteOrderMarkIsNotTakenForUtf16() {
        byte[] utf32 = ("\ufeff" + "héllo").getBytes(Charset.forName("UTF-32LE"));

        assertThat(map("text/plain", utf32).item().getContentItem().getContent()
                .getContentValue()).isEqualTo("héllo");
    }

    @Test
    void aByteOrderMarkWinsOverTheDeclaredCharset() {
        byte[] utf8 = ("\ufeff" + "héllo").getBytes(StandardCharsets.UTF_8);

        assertThat(map("text/plain; charset=iso-8859-1", utf8).item().getContentItem()
                .getContent().getContentValue()).isEqualTo("héllo");
    }

    @Test
    void aQuotedCharsetParameterIsRead() {
        assertThat(map("text/plain; format=flowed; charset=\"ISO-8859-1\"",
                "héllo".getBytes(StandardCharsets.ISO_8859_1)).item().getContentItem()
                .getContent().getContentValue()).isEqualTo("héllo");
    }

    @Test
    void anAsciiReadableDeclarationOfUtf16IsNotBelieved() {
        byte[] html = "<meta charset=\"utf-16\"><p>héllo</p>".getBytes(StandardCharsets.UTF_8);

        assertThat(map("text/html", html).item().getContentItem().getContent()
                .getContentValue()).isEqualTo("<meta charset=\"utf-16\"><p>héllo</p>");
    }

    @Test
    void richTextMarkupAndFilesWithoutAMediaTypeAreNotText() {
        byte[] rtf = "{\\rtf1\\ansi hello}".getBytes(StandardCharsets.US_ASCII);

        for (String mime : new String[] {"text/rtf", "application/rtf", "",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel", "image/png", "application/zip"}) {
            CrawlItemMapper.Mapped<CrawlItem> mapped = map(mime, rtf);
            assertThat(mapped.item().getContentItem().hasContent()).as(mime).isFalse();
            assertThat(mapped.omittedContent().orElseThrow().reason()).as(mime)
                    .isEqualTo(OmittedContent.Reason.NOT_TEXT);
        }
    }

    @Test
    void mediaTypesAreMatchedWithoutRegardToCase() {
        assertThat(map("Text/Plain; Charset=UTF-8", "hello".getBytes(StandardCharsets.UTF_8))
                .item().getContentItem().getContent().getContentValue()).isEqualTo("hello");
    }

    private static CrawlItemMapper.Mapped<CrawlItem> map(String mimeType, byte[] content) {
        return CrawlItemMapper.toCrawlItem(DriveItem.newBuilder()
                .setId("file-1")
                .setName("file")
                .setDriveId("drive-1")
                .setMimeType(mimeType)
                .setSize(content.length)
                .setContent(ByteString.copyFrom(content))
                .build(), true);
    }
}
