package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.ManagedChannel;
import io.grpc.stub.StreamObserver;
import microsoft.graph.connectors.contracts.grpc.ConnectorCrawlerServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.CrawlCheckpoint;
import microsoft.graph.connectors.contracts.grpc.CrawlItem;
import microsoft.graph.connectors.contracts.grpc.CrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.GetCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.GetIncrementalCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlItem;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.OperationResult;

import java.util.Iterator;
import java.util.Optional;

/**
 * GCA crawl streams: one {@code MicrosoftService.Sync} pass, mapped onto
 * {@link CrawlItem} / {@link IncrementalCrawlItem}.
 *
 * <p>Checkpoint {@code customMarkerData} is only ever a Sync resume cursor
 * (or empty): every checkpoint written while a pass runs repeats the cursor
 * the request carried, and the last bit carries the cursor the pass ended
 * with. So a crawl resumed from any checkpoint has a cursor Sync issued.
 *
 * <p>An incremental crawl hands the stored cursor back as
 * {@code SyncRequest.since_cursor}, which makes the pass a delta pass:
 * only changes since that cursor, deletes included. With no stored cursor
 * it is a full enumeration, which carries no deletes. A full crawl never
 * sends a cursor; the platform works out its deletes from what the crawl
 * did not return.
 *
 * <p>A Sync stream that ends without its terminal resume cursor was cut
 * short, so the crawl ends with an error status and no new checkpoint.
 */
public final class ConnectorCrawlerServiceImpl
        extends ConnectorCrawlerServiceGrpc.ConnectorCrawlerServiceImplBase {

    /** Creates the GCA service implementation. */
    public ConnectorCrawlerServiceImpl() {
    }

    private static final System.Logger LOG =
            System.getLogger(ConnectorCrawlerServiceImpl.class.getName());

    /** Why a pass whose stream had no terminal cursor fails. */
    static final String NO_RESUME_CURSOR = "Sync ended without its resume cursor";

    @Override
    public void getCrawlStream(GetCrawlStreamRequest request,
            StreamObserver<CrawlStreamBit> observer) {
        ConnectorCustomConfig config;
        try {
            config = ConnectionManagementServiceImpl.parse(request.getCustomConfiguration(),
                    request.getAuthenticationData());
        } catch (Throwable t) {
            observer.onNext(CrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.fromThrowable(t))
                    .build());
            observer.onCompleted();
            return;
        }
        ManagedChannel channel = MicrosoftServiceClients.channel(config);
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                    MicrosoftServiceClients.stub(channel);
            String stored = request.getCrawlProgressMarker().getCustomMarkerData();
            Iterator<SyncResponse> sync = stub.sync(syncRequest(config, ""));
            int emitted = 0;
            String next = "";
            while (sync.hasNext()) {
                SyncResponse event = sync.next();
                if (event.getEventCase() == SyncResponse.EventCase.CHANGE) {
                    Optional<CrawlItem> item = CrawlItemMapper.toCrawlItem(event.getChange());
                    if (item.isEmpty()) {
                        continue;
                    }
                    emitted++;
                    observer.onNext(CrawlStreamBit.newBuilder()
                            .setStatus(OperationStatuses.success())
                            .setCrawlItem(item.get())
                            .setCrawlProgressMarker(checkpoint(emitted, stored))
                            .build());
                } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                    next = event.getResumeCursor();
                }
            }
            if (next.isEmpty()) {
                throw new IllegalStateException(NO_RESUME_CURSOR);
            }
            observer.onNext(CrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.success("emitted " + emitted))
                    .setCrawlProgressMarker(checkpoint(emitted, next))
                    .build());
            observer.onCompleted();
        } catch (Throwable t) {
            LOG.log(System.Logger.Level.WARNING, "GetCrawlStream failed: {0}", t.toString());
            observer.onNext(CrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.fromThrowable(t))
                    .build());
            observer.onCompleted();
        } finally {
            MicrosoftServiceClients.shutdown(channel);
        }
    }

    @Override
    public void getIncrementalCrawlStream(GetIncrementalCrawlStreamRequest request,
            StreamObserver<IncrementalCrawlStreamBit> observer) {
        ConnectorCustomConfig config;
        try {
            config = ConnectionManagementServiceImpl.parse(request.getCustomConfiguration(),
                    request.getAuthenticationData());
        } catch (Throwable t) {
            observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.fromThrowable(t))
                    .build());
            observer.onCompleted();
            return;
        }
        ManagedChannel channel = MicrosoftServiceClients.channel(config);
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                    MicrosoftServiceClients.stub(channel);
            String stored = request.getCrawlProgressMarker().getCustomMarkerData();
            Iterator<SyncResponse> sync = stub.sync(syncRequest(config, stored));
            int emitted = 0;
            String next = "";
            while (sync.hasNext()) {
                SyncResponse event = sync.next();
                if (event.getEventCase() == SyncResponse.EventCase.CHANGE) {
                    MicrosoftChange change = event.getChange();
                    Optional<IncrementalCrawlItem> item =
                            CrawlItemMapper.toIncrementalCrawlItem(change);
                    if (item.isEmpty()) {
                        continue;
                    }
                    emitted++;
                    observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                            .setStatus(OperationStatuses.success())
                            .setCrawlItem(item.get())
                            .setCrawlProgressMarker(checkpoint(emitted, stored))
                            .build());
                } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                    next = event.getResumeCursor();
                }
            }
            if (next.isEmpty()) {
                throw new IllegalStateException(NO_RESUME_CURSOR);
            }
            observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.success("emitted " + emitted))
                    .setCrawlProgressMarker(checkpoint(emitted, next))
                    .build());
            observer.onCompleted();
        } catch (Throwable t) {
            LOG.log(System.Logger.Level.WARNING, "GetIncrementalCrawlStream failed: {0}",
                    t.toString());
            observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.of(OperationResult.DatasourceError,
                            t.getMessage()))
                    .build());
            observer.onCompleted();
        } finally {
            MicrosoftServiceClients.shutdown(channel);
        }
    }

    /** One pass over the configured scope; an empty cursor asks for a full enumeration. */
    private static SyncRequest syncRequest(ConnectorCustomConfig config, String sinceCursor) {
        return SyncRequest.newBuilder()
                .addAllDriveIds(config.driveIds())
                .setFolderPath(config.folderPath())
                .setIncludeContent(config.includeContent())
                .setSinceCursor(sinceCursor)
                .build();
    }

    /** A checkpoint whose marker is a Sync resume cursor, or empty for none. */
    private static CrawlCheckpoint checkpoint(int emitted, String resumeCursor) {
        return CrawlCheckpoint.newBuilder()
                .setPagenumber(1)
                .setBatchSize(emitted)
                .setCustomMarkerData(resumeCursor)
                .build();
    }
}
