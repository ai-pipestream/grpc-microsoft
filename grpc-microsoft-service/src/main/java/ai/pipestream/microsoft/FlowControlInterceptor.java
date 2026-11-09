package ai.pipestream.microsoft;

import io.grpc.Context;
import io.grpc.ForwardingServerCall;
import io.grpc.Metadata;
import io.grpc.ServerCall;
import io.grpc.ServerCallHandler;
import io.grpc.ServerInterceptor;
import io.grpc.Status;

import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.LockSupport;

/**
 * Outbound flow control for every server-streaming rpc: each message waits,
 * on the (virtual) thread that sends it, until the transport is ready for
 * it, so a slow client makes paging slow down instead of making the server
 * buffer whole result sets. A cancelled call (or its context) fails the
 * send with {@code CANCELLED}, which ends the handler's paging at once.
 *
 * <p>Readiness is polled with a short backoff rather than signalled: the
 * blocking handler runs inside the call's serialized listener callback,
 * so {@code onReady} notifications cannot be delivered until it returns,
 * while {@link ServerCall#isReady()} and {@link ServerCall#isCancelled()}
 * read transport state directly and are safe from any thread. Unary rpcs
 * are passed through untouched.</p>
 */
final class FlowControlInterceptor implements ServerInterceptor {

    private static final long MIN_PARK_NANOS = TimeUnit.MICROSECONDS.toNanos(200);
    private static final long MAX_PARK_NANOS = TimeUnit.MILLISECONDS.toNanos(50);

    @Override
    public <Q, R> ServerCall.Listener<Q> interceptCall(ServerCall<Q, R> call, Metadata headers,
            ServerCallHandler<Q, R> next) {
        if (call.getMethodDescriptor().getType().serverSendsOneMessage()) {
            return next.startCall(call, headers);
        }
        // Captured here: senders may be crawler worker threads with no context attached.
        Context context = Context.current();
        return next.startCall(new ForwardingServerCall.SimpleForwardingServerCall<>(call) {
            @Override
            public void sendMessage(R message) {
                awaitReady(this, context);
                super.sendMessage(message);
            }
        }, headers);
    }

    /** Parks until {@code call} is ready; CANCELLED when the call or its context is cancelled. */
    static void awaitReady(ServerCall<?, ?> call, Context context) {
        long park = MIN_PARK_NANOS;
        while (true) {
            if (call.isCancelled() || context.isCancelled()) {
                throw Status.CANCELLED.withDescription("the client cancelled the call")
                        .asRuntimeException();
            }
            if (call.isReady()) {
                return;
            }
            LockSupport.parkNanos(park);
            if (Thread.interrupted()) {
                Thread.currentThread().interrupt();
                throw Status.CANCELLED.withDescription("interrupted while waiting for the client")
                        .asRuntimeException();
            }
            park = Math.min(park * 2, MAX_PARK_NANOS);
        }
    }
}
