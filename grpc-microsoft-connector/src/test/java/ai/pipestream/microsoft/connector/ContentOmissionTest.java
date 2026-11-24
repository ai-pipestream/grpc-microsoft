package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ContentState;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import com.google.protobuf.ByteString;
import com.google.protobuf.Timestamp;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import microsoft.graph.connectors.contracts.grpc.Content;
import microsoft.graph.connectors.contracts.grpc.ConnectorCrawlerServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.ContentItem;
import microsoft.graph.connectors.contracts.grpc.CrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.CustomConfiguration;
import microsoft.graph.connectors.contracts.grpc.GetCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.GetIncrementalCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.OperationResult;
import microsoft.graph.connectors.contracts.grpc.OperationStatus;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.net.InetSocketAddress;
import java.net.ServerSocket;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * With {@code includeContent} on, a crawl item either carries the file's
 * text, correctly decoded, or says on its own status that the content was
 * left out and why. It never reports plain success for a file whose content
 * it dropped, and it never sends bytes that are not text as text.
 */
class ContentOmissionTest {

    private static final String DOCX =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private Server microsoftServer;
    private Server connectorServer;
    private ManagedChannel connectorChannel;
    private FakeMicrosoftService microsoft;
    private ConnectorCrawlerServiceGrpc.ConnectorCrawlerServiceBlockingStub crawler;
    private int port;

    @BeforeEach
    void startStack() throws Exception {
        microsoft = new FakeMicrosoftService();
        try (ServerSocket socket = new ServerSocket(0)) {
            port = socket.getLocalPort();
        }
        microsoftServer = io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder
                .forAddress(new InetSocketAddress("127.0.0.1", port))
                .addService(microsoft)
                .build()
                .start();
        String name = InProcessServerBuilder.generateName();
        connectorServer = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(new ConnectorCrawlerServiceImpl())
                .build()
                .start();
        connectorChannel = InProcessChannelBuilder.forName(name).directExecutor().build();
        crawler = ConnectorCrawlerServiceGrpc.newBlockingStub(connectorChannel);
    }

    @AfterEach
    void stopStack() throws InterruptedException {
        if (connectorChannel != null) {
            connectorChannel.shutdownNow();
        }
        if (connectorServer != null) {
            connectorServer.shutdownNow();
            connectorServer.awaitTermination(5, TimeUnit.SECONDS);
        }
        if (microsoftServer != null) {
            microsoftServer.shutdownNow();
            microsoftServer.awaitTermination(5, TimeUnit.SECONDS);
        }
    }

    @Test
    void anOfficeDocumentIsNotSentAsText() {
        // A docx is a zip. Its media type contains "xml" but its bytes are not text.
        byte[] zip = {'P', 'K', 3, 4, 20, 0, 6, 0, 8, 0, (byte) 0x9d, (byte) 0xe2, (byte) 0xff};
        microsoft.changes.add(file("doc-1", "plan.docx", DOCX, zip));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getCrawlItem().getContentItem().hasContent())
                .as("zip bytes sent as the document's text").isFalse();
        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage()).contains("drive-1/doc-1", DOCX);
    }

    @Test
    void aFileWhoseContentIsLeftOutSaysSoOnItsOwnStatus() {
        microsoft.changes.add(file("pdf-1", "report.pdf", "application/pdf",
                "%PDF-1.7".getBytes(StandardCharsets.US_ASCII)));
        microsoft.changes.add(file("txt-1", "notes.txt", "text/plain",
                "hello".getBytes(StandardCharsets.UTF_8)));

        List<CrawlStreamBit> bits = fullCrawl(true);

        // The item is still indexed by its properties.
        assertThat(bits.get(0).getCrawlItem().getContentItem().getPropertyValues().getValuesMap())
                .containsKey(DataSourceSchemas.TITLE);
        assertThat(bits.get(0).getCrawlItem().getContentItem().hasContent()).isFalse();
        assertThat(bits.get(0).getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bits.get(0).getStatus().getStatusMessage())
                .contains("drive-1/pdf-1", "application/pdf");
        // The text file next to it is a plain success with its text.
        assertThat(bits.get(1).getStatus().getResult()).isEqualTo(OperationResult.Success);
        assertThat(bits.get(1).getCrawlItem().getContentItem().getContent().getContentValue())
                .isEqualTo("hello");
        // The crawl ends normally and counts what it left out.
        CrawlStreamBit last = bits.get(bits.size() - 1);
        assertThat(last.getStatus().getResult()).isEqualTo(OperationResult.Success);
        assertThat(last.getCrawlProgressMarker().getCustomMarkerData()).isEqualTo("cursor-1");
        assertThat(last.getStatus().getStatusMessage()).contains("emitted 2", "1");
    }

    @Test
    void anIncrementalCrawlMarksLeftOutContentTheSameWay() {
        microsoft.changes.add(file("pdf-1", "report.pdf", "application/pdf",
                "%PDF-1.7".getBytes(StandardCharsets.US_ASCII)));

        List<IncrementalCrawlStreamBit> bits = new ArrayList<>();
        crawler.getIncrementalCrawlStream(GetIncrementalCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config(true))
                .build()).forEachRemaining(bits::add);

        assertThat(bits.get(0).getCrawlItem().getContentItem().hasContent()).isFalse();
        assertThat(bits.get(0).getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bits.get(0).getStatus().getStatusMessage())
                .contains("drive-1/pdf-1", "application/pdf");
    }

    @Test
    void aFileTooLargeToSendIsMarkedAndNeverDownloaded() {
        // The crawl tells Sync not to inline what could never be sent on.
        microsoft.changes.add(file("big-1", "big.log", "text/plain", null,
                CrawlItemMapper.CONTENT_CAP_BYTES + 1L));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(microsoft.syncRequests.get(0).getIncludeContent()).isTrue();
        assertThat(microsoft.syncRequests.get(0).getMaxInlineBytes())
                .isEqualTo(CrawlItemMapper.CONTENT_CAP_BYTES);
        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage())
                .contains("drive-1/big-1", String.valueOf(CrawlItemMapper.CONTENT_CAP_BYTES));
    }

    @Test
    void aTextFileTheServiceDidNotInlineIsMarked() {
        microsoft.changes.add(file("txt-1", "notes.txt", "text/plain", null, 12));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage()).contains("drive-1/txt-1");
    }

    @Test
    void nothingIsMarkedWhenTheConnectionDoesNotAskForContent() {
        microsoft.changes.add(file("pdf-1", "report.pdf", "application/pdf", null, 4096));
        microsoft.changes.add(file("txt-1", "notes.txt", "text/plain", null, 12));

        List<CrawlStreamBit> bits = fullCrawl(false);

        assertThat(microsoft.syncRequests.get(0).getIncludeContent()).isFalse();
        assertThat(bits).extracting(bit -> bit.getStatus().getResult())
                .containsOnly(OperationResult.Success);
    }

    @Test
    void foldersAndEmptyFilesHaveNoContentToLeaveOut() {
        microsoft.changes.add(change(DriveItem.newBuilder().setId("dir-1").setName("Shared")
                .setDriveId("drive-1").setFolder(true).build()));
        microsoft.changes.add(file("empty-1", "empty.txt", "text/plain", null, 0));

        List<CrawlStreamBit> bits = fullCrawl(true);

        assertThat(bits).extracting(bit -> bit.getStatus().getResult())
                .containsOnly(OperationResult.Success);
    }

    @Test
    void aUtf16FileIsDecodedByItsByteOrderMark() {
        microsoft.changes.add(file("txt-1", "notes.txt", "text/plain",
                withBom(new byte[] {(byte) 0xff, (byte) 0xfe}, "Größe: 5 €",
                        StandardCharsets.UTF_16LE)));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(content(bit).getContentValue()).isEqualTo("Größe: 5 €");
        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.Success);
    }

    @Test
    void aUtf8ByteOrderMarkIsNotPartOfTheText() {
        microsoft.changes.add(file("csv-1", "people.csv", "text/csv",
                withBom(new byte[] {(byte) 0xef, (byte) 0xbb, (byte) 0xbf}, "name,city",
                        StandardCharsets.UTF_8)));

        assertThat(content(fullCrawl(true).get(0)).getContentValue()).isEqualTo("name,city");
    }

    @Test
    void theCharsetTheMediaTypeDeclaresIsUsed() {
        microsoft.changes.add(file("txt-1", "menu.txt", "text/plain; charset=windows-1252",
                "café, 5 €".getBytes(Charset.forName("windows-1252"))));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(content(bit).getContentValue()).isEqualTo("café, 5 €");
        assertThat(content(bit).getContentType()).isEqualTo(Content.ContentType.Text);
    }

    @Test
    void anHtmlPageIsDecodedWithTheCharsetItDeclares() {
        String html = "<html><head><meta charset=\"iso-8859-1\"></head><body>Übersicht</body></html>";
        microsoft.changes.add(file("page-1", "index.html", "text/html",
                html.getBytes(StandardCharsets.ISO_8859_1)));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(content(bit).getContentValue()).isEqualTo(html);
        assertThat(content(bit).getContentType()).isEqualTo(Content.ContentType.Html);
    }

    @Test
    void anXmlFileIsDecodedWithTheEncodingItDeclares() {
        String xml = "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><note>Grüße</note>";
        microsoft.changes.add(file("xml-1", "note.xml", "application/xml",
                xml.getBytes(StandardCharsets.ISO_8859_1)));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(content(bit).getContentValue()).isEqualTo(xml);
        assertThat(content(bit).getContentType()).isEqualTo(Content.ContentType.Text);
    }

    @Test
    void textThatIsNotValidInItsCharsetIsLeftOutNotSentAsGarbage() {
        // Latin-1 bytes with no declared charset: not valid UTF-8.
        microsoft.changes.add(file("txt-1", "menu.txt", "text/plain",
                "café".getBytes(StandardCharsets.ISO_8859_1)));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getCrawlItem().getContentItem().hasContent())
                .as("content with replacement characters: %s",
                        bit.getCrawlItem().getContentItem().getContent().getContentValue())
                .isFalse();
        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage()).contains("drive-1/txt-1", "UTF-8");
    }

    @Test
    void aCharsetThisJvmDoesNotKnowLeavesTheContentOut() {
        microsoft.changes.add(file("txt-1", "notes.txt", "text/plain; charset=x-no-such-charset",
                "hello".getBytes(StandardCharsets.UTF_8)));

        OperationStatus status = fullCrawl(true).get(0).getStatus();

        assertThat(status.getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(status.getStatusMessage()).contains("x-no-such-charset");
    }

    @Test
    void structuredTextTypesAreSentAsText() {
        microsoft.changes.add(file("json-1", "a.json", "application/json",
                "{\"a\":1}".getBytes(StandardCharsets.UTF_8)));
        microsoft.changes.add(file("feed-1", "feed.atom", "application/atom+xml",
                "<feed/>".getBytes(StandardCharsets.UTF_8)));
        microsoft.changes.add(file("page-1", "page.xhtml", "application/xhtml+xml",
                "<html/>".getBytes(StandardCharsets.UTF_8)));

        List<CrawlStreamBit> bits = fullCrawl(true);

        assertThat(content(bits.get(0)).getContentValue()).isEqualTo("{\"a\":1}");
        assertThat(content(bits.get(0)).getContentType()).isEqualTo(Content.ContentType.Text);
        assertThat(content(bits.get(1)).getContentValue()).isEqualTo("<feed/>");
        assertThat(content(bits.get(2)).getContentType()).isEqualTo(Content.ContentType.Html);
        assertThat(bits).extracting(bit -> bit.getStatus().getResult())
                .containsOnly(OperationResult.Success);
    }

    private static Content content(CrawlStreamBit bit) {
        ContentItem item = bit.getCrawlItem().getContentItem();
        assertThat(item.hasContent()).as("content of %s: %s", bit.getCrawlItem().getItemId(),
                bit.getStatus().getStatusMessage()).isTrue();
        return item.getContent();
    }

    private static byte[] withBom(byte[] bom, String text, Charset charset) {
        byte[] encoded = text.getBytes(charset);
        byte[] out = new byte[bom.length + encoded.length];
        System.arraycopy(bom, 0, out, 0, bom.length);
        System.arraycopy(encoded, 0, out, bom.length, encoded.length);
        return out;
    }

    @Test
    void aFileTheServiceFoundOverItsLimitIsMarkedTooLargeNotJustMissing() {
        // Listed as 12 bytes, so nothing here could tell it was too large.
        microsoft.changes.add(change(DriveItem.newBuilder()
                .setId("txt-1").setName("notes.txt").setDriveId("drive-1")
                .setMimeType("text/plain").setSize(12)
                .setContentState(ContentState.CONTENT_STATE_OMITTED_OVER_LIMIT).build()));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage())
                .contains("drive-1/txt-1", "TOO_LARGE", "inline limit");
    }

    @Test
    void aFileFlaggedAsMalwareSaysSo() {
        microsoft.changes.add(change(DriveItem.newBuilder()
                .setId("bad-1").setName("notes.txt").setDriveId("drive-1")
                .setMimeType("text/plain").setSize(12)
                .setContentState(ContentState.CONTENT_STATE_OMITTED_MALWARE).build()));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getCrawlItem().getContentItem().hasContent()).isFalse();
        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.PartialSuccess);
        assertThat(bit.getStatus().getStatusMessage()).contains("drive-1/bad-1", "MALWARE");
    }

    @Test
    void aPackageHasNoContentToLeaveOut() {
        microsoft.changes.add(change(DriveItem.newBuilder()
                .setId("nb-1").setName("Notebook").setDriveId("drive-1").setSize(4096)
                .setContentState(ContentState.CONTENT_STATE_OMITTED_NOT_A_FILE).build()));

        CrawlStreamBit bit = fullCrawl(true).get(0);

        assertThat(bit.getStatus().getResult()).isEqualTo(OperationResult.Success);
    }

    private List<CrawlStreamBit> fullCrawl(boolean includeContent) {
        List<CrawlStreamBit> bits = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config(includeContent))
                .build()).forEachRemaining(bits::add);
        return bits;
    }

    private CustomConfiguration config(boolean includeContent) {
        return CustomConfiguration.newBuilder()
                .setConfiguration("{\"target\":\"127.0.0.1:" + port + "\",\"plaintext\":true,"
                        + "\"includeContent\":" + includeContent + "}")
                .build();
    }

    private static MicrosoftChange file(String id, String name, String mimeType, byte[] content) {
        return file(id, name, mimeType, content, content.length);
    }

    /** A file change; {@code content} is {@code null} when Sync did not inline it. */
    private static MicrosoftChange file(String id, String name, String mimeType, byte[] content,
            long size) {
        DriveItem.Builder item = DriveItem.newBuilder()
                .setId(id)
                .setName(name)
                .setDriveId("drive-1")
                .setMimeType(mimeType)
                .setSize(size);
        if (content != null) {
            item.setContent(ByteString.copyFrom(content));
        }
        return change(item.build());
    }

    private static MicrosoftChange change(DriveItem item) {
        return MicrosoftChange.newBuilder()
                .setChangeId("c-" + item.getId())
                .setOperation(ChangeOperation.CHANGE_OPERATION_UPSERT)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(item.getId())
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1_700_000_200))
                        .setDriveItem(item))
                .build();
    }
}
