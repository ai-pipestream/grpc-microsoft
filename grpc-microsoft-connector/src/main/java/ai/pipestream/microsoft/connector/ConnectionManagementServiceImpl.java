package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.ManagedChannel;
import io.grpc.stub.StreamObserver;
import microsoft.graph.connectors.contracts.grpc.AuthenticationData;
import microsoft.graph.connectors.contracts.grpc.ConnectionManagementServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.CustomConfiguration;
import microsoft.graph.connectors.contracts.grpc.GetDataSourceSchemaRequest;
import microsoft.graph.connectors.contracts.grpc.GetDataSourceSchemaResponse;
import microsoft.graph.connectors.contracts.grpc.OperationResult;
import microsoft.graph.connectors.contracts.grpc.ValidateAuthenticationRequest;
import microsoft.graph.connectors.contracts.grpc.ValidateAuthenticationResponse;
import microsoft.graph.connectors.contracts.grpc.ValidateCustomConfigurationRequest;
import microsoft.graph.connectors.contracts.grpc.ValidateCustomConfigurationResponse;

/**
 * GCA connection-creation APIs. Authentication is a {@code GetMe} against
 * this repo's MicrosoftService; the schema is the drive-item property list.
 */
public final class ConnectionManagementServiceImpl
        extends ConnectionManagementServiceGrpc.ConnectionManagementServiceImplBase {

    /** Creates the GCA service implementation. */
    public ConnectionManagementServiceImpl() {
    }

    @Override
    public void validateAuthentication(ValidateAuthenticationRequest request,
            StreamObserver<ValidateAuthenticationResponse> observer) {
        try {
            ConnectorCustomConfig config = parse(CustomConfiguration.getDefaultInstance(),
                    request.getAuthenticationData());
            GetMeResponse me = getMe(config);
            observer.onNext(ValidateAuthenticationResponse.newBuilder()
                    .setStatus(OperationStatuses.success("signed in as "
                            + me.getUser().getUserPrincipalName()))
                    .build());
            observer.onCompleted();
        } catch (Throwable t) {
            observer.onNext(ValidateAuthenticationResponse.newBuilder()
                    .setStatus(OperationStatuses.fromThrowable(t))
                    .build());
            observer.onCompleted();
        }
    }

    @Override
    public void validateCustomConfiguration(ValidateCustomConfigurationRequest request,
            StreamObserver<ValidateCustomConfigurationResponse> observer) {
        try {
            parse(request.getCustomConfiguration(), request.getAuthenticationData());
            observer.onNext(ValidateCustomConfigurationResponse.newBuilder()
                    .setStatus(OperationStatuses.success())
                    .build());
            observer.onCompleted();
        } catch (Throwable t) {
            observer.onNext(ValidateCustomConfigurationResponse.newBuilder()
                    .setStatus(OperationStatuses.of(OperationResult.ValidationFailure,
                            t.getMessage()))
                    .build());
            observer.onCompleted();
        }
    }

    @Override
    public void getDataSourceSchema(GetDataSourceSchemaRequest request,
            StreamObserver<GetDataSourceSchemaResponse> observer) {
        try {
            parse(request.getCustomConfiguration(), request.getAuthenticationData());
            observer.onNext(GetDataSourceSchemaResponse.newBuilder()
                    .setStatus(OperationStatuses.success())
                    .setDataSourceSchema(DataSourceSchemas.driveItem())
                    .build());
            observer.onCompleted();
        } catch (Throwable t) {
            observer.onNext(GetDataSourceSchemaResponse.newBuilder()
                    .setStatus(OperationStatuses.fromThrowable(t))
                    .build());
            observer.onCompleted();
        }
    }

    static ConnectorCustomConfig parse(CustomConfiguration custom, AuthenticationData auth) {
        String json = custom == null ? "" : custom.getConfiguration();
        String datasourceUrl = auth == null ? "" : auth.getDatasourceUrl();
        return ConnectorCustomConfig.parse(json, datasourceUrl);
    }

    private static GetMeResponse getMe(ConnectorCustomConfig config) {
        ManagedChannel channel = MicrosoftServiceClients.channel(config);
        try {
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub =
                    MicrosoftServiceClients.stub(channel);
            return stub.getMe(GetMeRequest.getDefaultInstance());
        } finally {
            MicrosoftServiceClients.shutdown(channel);
        }
    }
}
