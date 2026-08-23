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
 * {@link CrawlItem} / {@link IncrementalCrawlItem}. Checkpoint
 * {@code customMarkerData} is the last emitted item id, then the Sync
 * resume cursor.
 */
public final class ConnectorCrawlerServiceImpl
        extends ConnectorCrawlerServiceGrpc.ConnectorCrawlerServiceImplBase {

    /** Creates the GCA service implementation. */
    public ConnectorCrawlerServiceImpl() {
    }

    private static final System.Logger LOG =
            System.getLogger(ConnectorCrawlerServiceImpl.class.getName());

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
            Iterator<SyncResponse> sync = stub.sync(syncRequest(config));
            int emitted = 0;
            String marker = request.getCrawlProgressMarker().getCustomMarkerData();
            while (sync.hasNext()) {
                SyncResponse event = sync.next();
                if (event.getEventCase() == SyncResponse.EventCase.CHANGE) {
                    Optional<CrawlItem> item = CrawlItemMapper.toCrawlItem(event.getChange());
                    if (item.isEmpty()) {
                        continue;
                    }
                    emitted++;
                    marker = item.get().getItemId();
                    observer.onNext(CrawlStreamBit.newBuilder()
                            .setStatus(OperationStatuses.success())
                            .setCrawlItem(item.get())
                            .setCrawlProgressMarker(checkpoint(emitted, marker))
                            .build());
                } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                    marker = event.getResumeCursor();
                }
            }
            observer.onNext(CrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.success("emitted " + emitted))
                    .setCrawlProgressMarker(checkpoint(emitted, marker))
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
            Iterator<SyncResponse> sync = stub.sync(syncRequest(config));
            int emitted = 0;
            String marker = request.getCrawlProgressMarker().getCustomMarkerData();
            if (marker.isEmpty() && request.hasPreviousCrawlStartTimeInUtc()) {
                marker = request.getPreviousCrawlStartTimeInUtc().toString();
            }
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
                    marker = item.get().getItemId();
                    observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                            .setStatus(OperationStatuses.success())
                            .setCrawlItem(item.get())
                            .setCrawlProgressMarker(checkpoint(emitted, marker))
                            .build());
                } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                    marker = event.getResumeCursor();
                }
            }
            observer.onNext(IncrementalCrawlStreamBit.newBuilder()
                    .setStatus(OperationStatuses.success("emitted " + emitted))
                    .setCrawlProgressMarker(checkpoint(emitted, marker))
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

    private static SyncRequest syncRequest(ConnectorCustomConfig config) {
        return SyncRequest.newBuilder()
                .addAllDriveIds(config.driveIds())
                .setFolderPath(config.folderPath())
                .setIncludeContent(config.includeContent())
                .build();
    }

    private static CrawlCheckpoint checkpoint(int emitted, String marker) {
        return CrawlCheckpoint.newBuilder()
                .setPagenumber(1)
                .setBatchSize(emitted)
                .setCustomMarkerData(marker == null ? "" : marker)
                .build();
    }
}
