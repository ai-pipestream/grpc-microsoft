package ai.pipestream.microsoft.connector;

import io.grpc.stub.StreamObserver;
import microsoft.graph.connectors.contracts.grpc.AuthenticationData;
import microsoft.graph.connectors.contracts.grpc.ConnectorOAuthServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.OAuth2ClientCredentialResponse;
import microsoft.graph.connectors.contracts.grpc.RefreshAccessTokenRequest;
import microsoft.graph.connectors.contracts.grpc.RefreshAccessTokenResponse;

/**
 * GCA OAuth refresh. Graph tokens are owned by {@code MicrosoftService};
 * this adapter echoes any token GCA already holds and reports success so
 * the agent does not fail the connection.
 */
public final class ConnectorOAuthServiceImpl
        extends ConnectorOAuthServiceGrpc.ConnectorOAuthServiceImplBase {

    @Override
    public void refreshAccessToken(RefreshAccessTokenRequest request,
            StreamObserver<RefreshAccessTokenResponse> observer) {
        OAuth2ClientCredentialResponse existing = existing(request.getAuthenticationData());
        observer.onNext(RefreshAccessTokenResponse.newBuilder()
                .setStatus(OperationStatuses.success(
                        "MicrosoftService owns Graph tokens; nothing to refresh here"))
                .setRefreshedCredentialData(existing)
                .build());
        observer.onCompleted();
    }

    private static OAuth2ClientCredentialResponse existing(AuthenticationData data) {
        if (data.hasOAuth2ClientCredential()
                && data.getOAuth2ClientCredential().hasOAuth2ClientCredentialResponse()) {
            return data.getOAuth2ClientCredential().getOAuth2ClientCredentialResponse();
        }
        return OAuth2ClientCredentialResponse.getDefaultInstance();
    }
}
