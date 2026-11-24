package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.microsoft.v1.WatchChangesRequest;
import ai.pipestream.microsoft.v1.WatchChangesResponse;
import com.google.protobuf.Any;
import com.google.rpc.RetryInfo;
import io.grpc.Context;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.protobuf.StatusProto;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.time.Clock;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** WatchChanges keeps the stream up across a throttled or unavailable pass. */
class MicrosoftChangeRetryTest {

    private final List<Duration> slept = new CopyOnWriteArrayList<>();
    private Server server;
    private ManagedChannel channel;

    @AfterEach
    void stop() throws InterruptedException {
        if (channel != null) {
            channel.shutdownNow();
        }
        if (server != null) {
            server.shutdownNow();
            server.awaitTermination(5, TimeUnit.SECONDS);
        }
    }

    @Test
    void aThrottledPassRunsAgainFromTheSameCursorAfterRetryAfter() throws Exception {
        AtomicInteger calls = new AtomicInteger();
        List<String> cursors = new CopyOnWriteArrayList<>();
        MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub stub = start((request, observer) -> {
            cursors.add(request.getSinceCursor());
            if (calls.incrementAndGet() == 1) {
                observer.onError(StatusProto.toStatusRuntimeException(com.google.rpc.Status.newBuilder()
                        .setCode(Status.Code.RESOURCE_EXHAUSTED.value()).setMessage("throttled")
                        .addDetails(Any.pack(RetryInfo.newBuilder().setRetryDelay(
                                com.google.protobuf.Duration.newBuilder().setSeconds(3)).build()))
                        .build()));
                return;
            }
            observer.onNext(SyncResponse.newBuilder().setChange(MicrosoftChange.newBuilder()
                    .setChangeId("c-1").setSource(ChangeSource.CHANGE_SOURCE_DELTA)).build());
            observer.onNext(SyncResponse.newBuilder().setResumeCursor("cursor-2").build());
            observer.onCompleted();
        });
        try (Context.CancellableContext watch = Context.current().withCancellation()) {
            List<WatchChangesResponse> seen = watch.call(() -> {
                Iterator<WatchChangesResponse> stream = stub.watchChanges(WatchChangesRequest.newBuilder()
                        .setResumeCursor("cursor-1")
                        .setPollInterval(com.google.protobuf.Duration.newBuilder().setSeconds(3600))
                        .build());
                List<WatchChangesResponse> out = new ArrayList<>();
                while (stream.hasNext()) {
                    WatchChangesResponse next = stream.next();
                    out.add(next);
                    if (next.hasPassCompleted()) {
                        break;
                    }
                }
                return out;
            });
            watch.cancel(null);
            assertThat(cursors).containsExactly("cursor-1", "cursor-1");
            assertThat(seen).hasSize(2);
            assertThat(seen.get(0).getChange().getChangeId()).isEqualTo("c-1");
            assertThat(seen.get(1).getPassCompleted().getChangeCount()).isEqualTo(1);
            assertThat(seen.get(1).getResumeCursor()).isEqualTo("cursor-2");
            assertThat(slept.stream().reduce(Duration.ZERO, Duration::plus)).isEqualTo(Duration.ofSeconds(3));
        }
    }

    @Test
    void aPermissionErrorEndsTheStream() throws Exception {
        MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub stub = start((request, observer) ->
                observer.onError(Status.PERMISSION_DENIED.withDescription("no grant").asRuntimeException()));
        Iterator<WatchChangesResponse> stream = stub.watchChanges(WatchChangesRequest.getDefaultInstance());
        assertThatThrownBy(stream::hasNext).isInstanceOfSatisfying(StatusRuntimeException.class,
                e -> assertThat(e.getStatus().getCode()).isEqualTo(Status.Code.PERMISSION_DENIED));
        assertThat(slept).isEmpty();
    }

    @Test
    void retryWaitBacksOffWithoutRetryAfterAndStaysUnderTheCap() {
        StatusRuntimeException plain = Status.UNAVAILABLE.asRuntimeException();
        assertThat(MicrosoftChangeGrpcService.retryWait(plain, 1)).isEqualTo(Duration.ofSeconds(2));
        assertThat(MicrosoftChangeGrpcService.retryWait(plain, 3)).isEqualTo(Duration.ofSeconds(8));
        assertThat(MicrosoftChangeGrpcService.retryWait(plain, 40))
                .isEqualTo(MicrosoftChangeGrpcService.MAX_RETRY_WAIT);
    }

    @Test
    void aMissedLifecycleNotificationStartsAPass() {
        assertThat(MicrosoftChangeGrpcService.startsPass(ChangeNotification.getDefaultInstance())).isTrue();
        assertThat(MicrosoftChangeGrpcService.startsPass(ChangeNotification.newBuilder()
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_MISSED).build())).isTrue();
        assertThat(MicrosoftChangeGrpcService.startsPass(ChangeNotification.newBuilder()
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED).build()))
                .isFalse();
        assertThat(MicrosoftChangeGrpcService.startsPass(ChangeNotification.newBuilder()
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED).build()))
                .isTrue();
    }

    private MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub start(
            MicrosoftChangeGrpcService.Syncer syncer) throws Exception {
        String name = InProcessServerBuilder.generateName();
        MicrosoftChangeGrpcService service = new MicrosoftChangeGrpcService(null, syncer,
                new NotificationBus(), false, "", Clock.systemUTC(), slept::add);
        server = InProcessServerBuilder.forName(name).addService(service).build().start();
        channel = InProcessChannelBuilder.forName(name).build();
        return MicrosoftChangeServiceGrpc.newBlockingStub(channel);
    }
}
