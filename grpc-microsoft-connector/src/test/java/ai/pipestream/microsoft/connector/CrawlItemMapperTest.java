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

import static org.assertj.core.api.Assertions.assertThat;

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

        CrawlItem crawl = CrawlItemMapper.toCrawlItem(item);

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
    void deleteBecomesIncrementalDeletedItem() {
        MicrosoftChange change = MicrosoftChange.newBuilder()
                .setChangeId("c1")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("gone")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1)))
                .build();

        IncrementalCrawlItem item = CrawlItemMapper.toIncrementalCrawlItem(change).orElseThrow();

        assertThat(item.getItemType()).isEqualTo(IncrementalCrawlItem.ItemType.DeletedItem);
        assertThat(item.getItemId()).isEqualTo("gone");
        assertThat(CrawlItemMapper.toCrawlItem(change)).isEmpty();
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

        assertThat(CrawlItemMapper.toCrawlItem(change)).isEmpty();
    }
}
