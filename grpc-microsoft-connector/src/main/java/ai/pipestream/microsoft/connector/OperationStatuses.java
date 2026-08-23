package ai.pipestream.microsoft.connector;

import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import microsoft.graph.connectors.contracts.grpc.OperationResult;
import microsoft.graph.connectors.contracts.grpc.OperationStatus;
import microsoft.graph.connectors.contracts.grpc.RetryDetails;

/**
 * Maps our MicrosoftService statuses onto the GCA {@link OperationStatus}
 * vocabulary.
 */
public final class OperationStatuses {

    private OperationStatuses() {
    }

    public static OperationStatus success() {
        return success("");
    }

    public static OperationStatus success(String message) {
        return OperationStatus.newBuilder()
                .setResult(OperationResult.Success)
                .setStatusMessage(message == null ? "" : message)
                .build();
    }

    public static OperationStatus of(OperationResult result, String message) {
        return OperationStatus.newBuilder()
                .setResult(result)
                .setStatusMessage(message == null ? "" : message)
                .build();
    }

    public static OperationStatus fromThrowable(Throwable t) {
        if (t instanceof IllegalArgumentException e) {
            return of(OperationResult.ValidationFailure, e.getMessage());
        }
        if (t instanceof StatusRuntimeException e) {
            return fromStatus(e.getStatus());
        }
        if (t instanceof InterruptedException e) {
            Thread.currentThread().interrupt();
            return of(OperationResult.Cancelled, e.getMessage());
        }
        return of(OperationResult.DatasourceError, String.valueOf(t.getMessage()));
    }

    public static OperationStatus fromStatus(Status status) {
        OperationResult result = switch (status.getCode()) {
            case UNAUTHENTICATED, PERMISSION_DENIED -> OperationResult.AuthenticationIssue;
            case INVALID_ARGUMENT, FAILED_PRECONDITION, OUT_OF_RANGE ->
                    OperationResult.ValidationFailure;
            case UNAVAILABLE -> OperationResult.NetworkError;
            case CANCELLED -> OperationResult.Cancelled;
            case DEADLINE_EXCEEDED, RESOURCE_EXHAUSTED -> OperationResult.DatasourceError;
            default -> OperationResult.DatasourceError;
        };
        OperationStatus.Builder builder = OperationStatus.newBuilder()
                .setResult(result)
                .setStatusMessage(status.getDescription() == null ? status.getCode().name()
                        : status.getDescription());
        if (result == OperationResult.NetworkError || status.getCode() == Status.Code.RESOURCE_EXHAUSTED) {
            builder.setRetryInfo(RetryDetails.newBuilder()
                    .setType(RetryDetails.RetryType.Standard)
                    .setNumberOfRetries(3)
                    .setPauseBetweenRetriesInMilliseconds(2_000));
        }
        return builder.build();
    }
}
