package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.WriteError;
import ai.pipestream.microsoft.v1.WriteErrorCode;
import com.google.protobuf.Any;
import com.google.protobuf.Duration;
import com.google.rpc.RetryInfo;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.protobuf.StatusProto;

import java.io.IOException;

/** Maps failures of a Graph call to gRPC statuses and typed write errors. */
final class GraphErrors {

    private GraphErrors() {
    }

    /**
     * The gRPC error for a failure, with {@code RetryInfo} when Graph sent
     * {@code Retry-After}.
     *
     * @param failure what a handler threw
     * @return the error to send to the client
     */
    static StatusRuntimeException toException(Throwable failure) {
        Throwable cause = unwrap(failure);
        if (cause instanceof StatusRuntimeException e) {
            return e;
        }
        Status status = toStatus(cause);
        long retryAfter = retryAfterSeconds(cause);
        if (retryAfter < 0) {
            return status.asRuntimeException();
        }
        return StatusProto.toStatusRuntimeException(com.google.rpc.Status.newBuilder()
                .setCode(status.getCode().value())
                .setMessage(String.valueOf(status.getDescription()))
                .addDetails(Any.pack(RetryInfo.newBuilder()
                        .setRetryDelay(Duration.newBuilder().setSeconds(retryAfter)).build()))
                .build());
    }

    /**
     * The typed error for a failed write.
     *
     * @param failure what the write threw
     * @return the error
     */
    static WriteError toWriteError(Throwable failure) {
        Throwable cause = unwrap(failure);
        Status status = cause instanceof StatusRuntimeException e ? e.getStatus() : toStatus(cause);
        WriteError.Builder error = WriteError.newBuilder()
                .setCode(code(status.getCode()))
                .setMessage(status.getDescription() == null ? "" : status.getDescription());
        long retryAfter = retryAfterSeconds(cause);
        if (retryAfter >= 0) {
            error.setRetryDelay(Duration.newBuilder().setSeconds(retryAfter));
        }
        return error.build();
    }

    private static Throwable unwrap(Throwable failure) {
        if (failure instanceof java.io.UncheckedIOException u && u.getCause() != null) {
            return u.getCause();
        }
        return failure;
    }

    private static long retryAfterSeconds(Throwable cause) {
        return cause instanceof GraphClient.GraphApiException e ? e.retryAfterSeconds() : -1L;
    }

    private static Status toStatus(Throwable cause) {
        if (cause instanceof GraphClient.GraphApiException e) {
            return (switch (e.status()) {
                case 400 -> Status.INVALID_ARGUMENT;
                case 401 -> Status.UNAUTHENTICATED;
                case 403 -> Status.PERMISSION_DENIED;
                case 404 -> Status.NOT_FOUND;
                case 409, 412 -> Status.ABORTED;
                case 410, 423, 424 -> Status.FAILED_PRECONDITION;
                case 416 -> Status.OUT_OF_RANGE;
                case 429, 507 -> Status.RESOURCE_EXHAUSTED;
                case 503, 504 -> Status.UNAVAILABLE;
                default -> e.code().equals("nameAlreadyExists") ? Status.ABORTED : Status.INTERNAL;
            }).withDescription(e.getMessage());
        }
        if (cause instanceof GraphPager.RepeatedLinkException e) {
            return Status.INTERNAL.withDescription(e.getMessage());
        }
        if (cause instanceof IllegalArgumentException e) {
            return Status.INVALID_ARGUMENT.withDescription(e.getMessage());
        }
        if (cause instanceof InterruptedException e) {
            Thread.currentThread().interrupt();
            return Status.CANCELLED.withDescription("interrupted: " + e.getMessage());
        }
        if (cause instanceof IOException e) {
            return Status.UNAVAILABLE.withDescription("Graph unreachable: " + e.getMessage());
        }
        if (cause instanceof IllegalStateException e) {
            return Status.UNAVAILABLE.withDescription(e.getMessage());
        }
        return Status.INTERNAL.withDescription(String.valueOf(cause.getMessage()));
    }

    private static WriteErrorCode code(Status.Code code) {
        return switch (code) {
            case INVALID_ARGUMENT, OUT_OF_RANGE -> WriteErrorCode.WRITE_ERROR_CODE_INVALID_ARGUMENT;
            case UNAUTHENTICATED -> WriteErrorCode.WRITE_ERROR_CODE_UNAUTHENTICATED;
            case PERMISSION_DENIED -> WriteErrorCode.WRITE_ERROR_CODE_PERMISSION_DENIED;
            case NOT_FOUND -> WriteErrorCode.WRITE_ERROR_CODE_NOT_FOUND;
            case ABORTED, ALREADY_EXISTS -> WriteErrorCode.WRITE_ERROR_CODE_ABORTED;
            case FAILED_PRECONDITION -> WriteErrorCode.WRITE_ERROR_CODE_FAILED_PRECONDITION;
            case RESOURCE_EXHAUSTED -> WriteErrorCode.WRITE_ERROR_CODE_RESOURCE_EXHAUSTED;
            case UNAVAILABLE, DEADLINE_EXCEEDED, CANCELLED -> WriteErrorCode.WRITE_ERROR_CODE_UNAVAILABLE;
            default -> WriteErrorCode.WRITE_ERROR_CODE_INTERNAL;
        };
    }
}
