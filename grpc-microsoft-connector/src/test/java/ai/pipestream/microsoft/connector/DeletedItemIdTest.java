package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.Deleted;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import com.google.protobuf.Timestamp;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import microsoft.graph.connectors.contracts.grpc.ConnectorCrawlerServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.CrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.CustomConfiguration;
import microsoft.graph.connectors.contracts.grpc.GetCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.GetIncrementalCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlItem;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.OperationResult;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.net.InetSocketAddress;
import java.net.ServerSocket;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * A delete names the item by the id its upsert was indexed under. An item
 * is indexed as {@code <drive id>/<item id>}; a delete sent under the bare
 * item id matches nothing, is acknowledged, and leaves the document in the
 * index.
 */
class DeletedItemIdTest {

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
    void aDeleteNamesTheItemItsUpsertIndexed() {
        // First pass: the file is indexed.
        microsoft.changes.add(FakeMicrosoftService.upsertFile("file-1", "notes.txt"));
        List<CrawlStreamBit> full = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(full::add);
        String indexedAs = full.get(0).getCrawlItem().getItemId();

        // Next pass: the file was deleted. The change is the one the crawler
        // emits for a delta tombstone: the entity id is the bare item id, and
        // the drive item carries its drive.
        microsoft.changes.clear();
        microsoft.changes.add(deleted("file-1", "drive-1"));
        List<IncrementalCrawlStreamBit> incremental = incrementalCrawl();

        IncrementalCrawlItem item = incremental.get(0).getCrawlItem();
        assertThat(item.getItemType()).isEqualTo(IncrementalCrawlItem.ItemType.DeletedItem);
        assertThat(item.getItemId())
                .as("the id a delete is sent under, for an item indexed as %s", indexedAs)
                .isEqualTo(indexedAs);
    }

    @Test
    void aDeleteThatCannotBeNamedFailsTheCrawlInsteadOfBeingAcknowledged() {
        microsoft.changes.add(MicrosoftChange.newBuilder()
                .setChangeId("d-gone")
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId("gone")
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1_700_000_300)))
                .build());

        List<IncrementalCrawlStreamBit> bits = incrementalCrawl();

        IncrementalCrawlStreamBit last = bits.get(bits.size() - 1);
        assertThat(last.getStatus().getResult()).isEqualTo(OperationResult.DatasourceError);
        assertThat(last.getStatus().getStatusMessage()).contains("d-gone", "gone");
        assertThat(last.getCrawlProgressMarker().getCustomMarkerData())
                .as("no new cursor is handed out past the delete").isEmpty();
        assertThat(bits).noneMatch(bit -> bit.getCrawlItem().getItemType()
                == IncrementalCrawlItem.ItemType.DeletedItem);
    }

    @Test
    void aFullCrawlStillCarriesNoDeletes() {
        microsoft.changes.add(deleted("file-1", "drive-1"));
        List<CrawlStreamBit> full = new ArrayList<>();

        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(full::add);

        assertThat(full).singleElement()
                .satisfies(bit -> assertThat(bit.hasCrawlItem()).isFalse());
    }

    private List<IncrementalCrawlStreamBit> incrementalCrawl() {
        List<IncrementalCrawlStreamBit> bits = new ArrayList<>();
        crawler.getIncrementalCrawlStream(GetIncrementalCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(bits::add);
        return bits;
    }

    private CustomConfiguration config() {
        return CustomConfiguration.newBuilder()
                .setConfiguration("{\"target\":\"127.0.0.1:" + port + "\",\"plaintext\":true}")
                .build();
    }

    /** A delete as {@code MicrosoftCrawler} emits it for a delta tombstone. */
    private static MicrosoftChange deleted(String id, String driveId) {
        return MicrosoftChange.newBuilder()
                .setChangeId("d-" + id)
                .setOperation(ChangeOperation.CHANGE_OPERATION_DELETE)
                .setEntity(MicrosoftEntity.newBuilder()
                        .setEntityId(id)
                        .setIngestedAt(Timestamp.newBuilder().setSeconds(1_700_000_300))
                        .setDriveItem(DriveItem.newBuilder()
                                .setId(id)
                                .setDriveId(driveId)
                                .setDeleted(Deleted.newBuilder().setState("deleted"))))
                .build();
    }
}
