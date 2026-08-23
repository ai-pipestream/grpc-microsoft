package ai.pipestream.microsoft.connector;

import io.grpc.stub.StreamObserver;
import microsoft.graph.connectors.contracts.grpc.ConnectorInfoServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.GetBasicConnectorInfoRequest;
import microsoft.graph.connectors.contracts.grpc.GetBasicConnectorInfoResponse;
import microsoft.graph.connectors.contracts.grpc.HealthCheckRequest;
import microsoft.graph.connectors.contracts.grpc.HealthCheckResponse;

/**
 * GCA {@code ConnectorInfoService}: stable connector id plus a no-op
 * health check. Connectivity to MicrosoftService is proven later by
 * {@code ValidateAuthentication}.
 */
public final class ConnectorInfoServiceImpl
        extends ConnectorInfoServiceGrpc.ConnectorInfoServiceImplBase {

    /** Creates the GCA service implementation. */
    public ConnectorInfoServiceImpl() {
    }

    @Override
    public void getBasicConnectorInfo(GetBasicConnectorInfoRequest request,
            StreamObserver<GetBasicConnectorInfoResponse> observer) {
        observer.onNext(GetBasicConnectorInfoResponse.newBuilder()
                .setConnectorId(ConnectorId.VALUE)
                .build());
        observer.onCompleted();
    }

    @Override
    public void healthCheck(HealthCheckRequest request,
            StreamObserver<HealthCheckResponse> observer) {
        observer.onNext(HealthCheckResponse.getDefaultInstance());
        observer.onCompleted();
    }
}
