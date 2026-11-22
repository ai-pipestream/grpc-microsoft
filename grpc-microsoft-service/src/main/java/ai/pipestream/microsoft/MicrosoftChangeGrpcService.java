package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.CreateSubscriptionRequest;
import ai.pipestream.microsoft.v1.CreateSubscriptionResponse;
import ai.pipestream.microsoft.v1.DeleteSubscriptionRequest;
import ai.pipestream.microsoft.v1.DeleteSubscriptionResponse;
import ai.pipestream.microsoft.v1.GetSubscriptionRequest;
import ai.pipestream.microsoft.v1.GetSubscriptionResponse;
import ai.pipestream.microsoft.v1.ListSubscriptionsRequest;
import ai.pipestream.microsoft.v1.ListSubscriptionsResponse;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.PassCompleted;
import ai.pipestream.microsoft.v1.PassTrigger;
import ai.pipestream.microsoft.v1.ReauthorizeSubscriptionRequest;
import ai.pipestream.microsoft.v1.ReauthorizeSubscriptionResponse;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import ai.pipestream.microsoft.v1.UpdateSubscriptionRequest;
import ai.pipestream.microsoft.v1.UpdateSubscriptionResponse;
import ai.pipestream.microsoft.v1.WatchChangesRequest;
import ai.pipestream.microsoft.v1.WatchChangesResponse;
import ai.pipestream.microsoft.v1.WatchNotificationsRequest;
import ai.pipestream.microsoft.v1.WatchNotificationsResponse;
import com.google.protobuf.Timestamp;
import io.grpc.Context;
import io.grpc.Status;
import io.grpc.stub.StreamObserver;

import java.time.Clock;
import java.time.Duration;
import java.time.Instant;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicLong;
import java.util.concurrent.atomic.AtomicReference;

/**
 * The continuous change feed over {@link MicrosoftGrpcService#sync} and the
 * Graph subscription rpcs.
 */
public final class MicrosoftChangeGrpcService
        implements MicrosoftChangeServiceGrpc.AsyncService, io.grpc.BindableService {

    /** Poll interval when a request leaves it unset. */
    static final Duration DEFAULT_POLL_INTERVAL = Duration.ofMinutes(5);
    /** Shortest wait between passes when a request leaves it unset. */
    static final Duration DEFAULT_MIN_INTERVAL = Duration.ofSeconds(5);
    /** How often a waiting stream checks whether its client has gone. */
    private static final Duration CANCEL_CHECK = Duration.ofSeconds(1);

    private final MicrosoftGrpcService microsoft;
    private final NotificationBus bus;
    private final boolean receiverOn;
    private final String clientState;
    private final Clock clock;

    /**
     * Creates the service.
     *
     * @param microsoft the read service whose Sync each pass runs
     * @param bus notifications from the receiver; never {@code null}
     * @param receiverOn whether a receiver feeds {@code bus}
     * @param clientState the clientState every new subscription carries; empty when the
     *        receiver is off
     * @param clock time source
     */
    MicrosoftChangeGrpcService(MicrosoftGrpcService microsoft, NotificationBus bus,
            boolean receiverOn, String clientState, Clock clock) {
        this.microsoft = microsoft;
        this.bus = bus;
        this.receiverOn = receiverOn;
        this.clientState = clientState == null ? "" : clientState;
        this.clock = clock;
    }

    /**
     * Binds the service behind {@link FlowControlInterceptor}, so a slow client
     * parks the stream instead of making the server buffer it.
     *
     * @return the intercepted service definition
     */
    @Override
    public io.grpc.ServerServiceDefinition bindService() {
        return io.grpc.ServerInterceptors.intercept(MicrosoftChangeServiceGrpc.bindService(this),
                new FlowControlInterceptor());
    }

    // ------------------------------------------------------------ change feed

    @Override
    public void watchChanges(WatchChangesRequest request,
            StreamObserver<WatchChangesResponse> observer) {
        Duration poll = request.hasPollInterval() ? duration(request.getPollInterval())
                : DEFAULT_POLL_INTERVAL;
        Duration min = request.hasMinInterval() ? duration(request.getMinInterval())
                : DEFAULT_MIN_INTERVAL;
        if (poll.isNegative() || poll.isZero() || min.isNegative()) {
            observer.onError(Status.INVALID_ARGUMENT
                    .withDescription("poll_interval must be positive and min_interval not negative")
                    .asRuntimeException());
            return;
        }
        Context context = Context.current();
        AtomicLong sequence = new AtomicLong();
        AtomicReference<String> cursor = new AtomicReference<>(request.getResumeCursor().isBlank()
                ? request.getSync().getSinceCursor() : request.getResumeCursor());
        try (NotificationBus.Listener listener = bus.listen()) {
            PassTrigger trigger = PassTrigger.PASS_TRIGGER_START;
            int notifications = 0;
            for (long pass = 1; !context.isCancelled(); pass++) {
                Instant started = clock.instant();
                long changes = runPass(request.getSync(), cursor, trigger, sequence, observer);
                synchronized (observer) {
                    observer.onNext(WatchChangesResponse.newBuilder()
                            .setSequence(sequence.incrementAndGet())
                            .setResumeCursor(cursor.get())
                            .setPassCompleted(PassCompleted.newBuilder().setPass(pass)
                                    .setTrigger(trigger).setNotificationCount(notifications)
                                    .setChangeCount(changes).setStartedAt(timestamp(started))
                                    .setCompletedAt(timestamp(clock.instant())))
                            .build());
                }
                Instant earliest = started.plus(min);
                Instant deadline = started.plus(poll);
                notifications = 0;
                trigger = PassTrigger.PASS_TRIGGER_TIMER;
                while (!context.isCancelled()) {
                    Instant now = clock.instant();
                    boolean due = !now.isBefore(deadline);
                    if (due || (notifications > 0 && !now.isBefore(earliest))) {
                        trigger = due && notifications == 0 ? PassTrigger.PASS_TRIGGER_TIMER
                                : PassTrigger.PASS_TRIGGER_NOTIFICATION;
                        break;
                    }
                    Instant wakeAt = notifications > 0 ? earliest : deadline;
                    long waitMillis = Math.max(1, Math.min(CANCEL_CHECK.toMillis(),
                            Duration.between(now, wakeAt).toMillis()));
                    ChangeNotification arrived = listener.queue().poll(waitMillis, TimeUnit.MILLISECONDS);
                    if (arrived != null && arrived.getLifecycleEvent()
                            == ai.pipestream.microsoft.v1.LifecycleEvent.LIFECYCLE_EVENT_UNSPECIFIED) {
                        notifications++;
                    }
                }
            }
            observer.onCompleted();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            observer.onError(Status.CANCELLED.withDescription("watch interrupted").asRuntimeException());
        } catch (Throwable t) {
            if (context.isCancelled()) {
                return;
            }
            observer.onError(GraphErrors.toException(t));
        }
    }

    /** Runs one Sync pass, forwarding its events; answers how many changes it emitted. */
    private long runPass(SyncRequest scope, AtomicReference<String> cursor, PassTrigger trigger,
            AtomicLong sequence, StreamObserver<WatchChangesResponse> observer) throws Throwable {
        AtomicLong changes = new AtomicLong();
        AtomicReference<Throwable> failure = new AtomicReference<>();
        AtomicReference<String> next = new AtomicReference<>();
        String resume = cursor.get();
        microsoft.sync(scope.toBuilder().setSinceCursor(resume).build(),
                new StreamObserver<>() {
                    @Override
                    public void onNext(SyncResponse value) {
                        WatchChangesResponse.Builder out = WatchChangesResponse.newBuilder()
                                .setResumeCursor(resume);
                        switch (value.getEventCase()) {
                            case CHANGE -> {
                                out.setChange(stamp(value.getChange(), trigger));
                                changes.incrementAndGet();
                            }
                            case SNAPSHOT -> out.setSnapshot(value.getSnapshot());
                            case RESUME_CURSOR -> {
                                next.set(value.getResumeCursor());
                                return;
                            }
                            default -> {
                                return;
                            }
                        }
                        synchronized (observer) {
                            observer.onNext(out.setSequence(sequence.incrementAndGet()).build());
                        }
                    }

                    @Override
                    public void onError(Throwable t) {
                        failure.set(t);
                    }

                    @Override
                    public void onCompleted() {
                        // The terminal resume cursor arrived in onNext.
                    }
                });
        if (failure.get() != null) {
            throw failure.get();
        }
        if (next.get() != null) {
            cursor.set(next.get());
        }
        return changes.get();
    }

    private static MicrosoftChange stamp(MicrosoftChange change, PassTrigger trigger) {
        if (trigger == PassTrigger.PASS_TRIGGER_NOTIFICATION
                && change.getSource() == ChangeSource.CHANGE_SOURCE_DELTA) {
            return change.toBuilder().setSource(ChangeSource.CHANGE_SOURCE_NOTIFICATION).build();
        }
        return change;
    }

    @Override
    public void watchNotifications(WatchNotificationsRequest request,
            StreamObserver<WatchNotificationsResponse> observer) {
        if (!receiverOn) {
            observer.onError(Status.FAILED_PRECONDITION.withDescription(
                    "the notification receiver is off; start the server with "
                            + NotificationReceiver.ENV_PORT).asRuntimeException());
            return;
        }
        Context context = Context.current();
        try (NotificationBus.Listener listener = bus.listen()) {
            while (!context.isCancelled()) {
                ChangeNotification next = listener.queue().poll(CANCEL_CHECK.toMillis(),
                        TimeUnit.MILLISECONDS);
                if (next != null && (request.getSubscriptionId().isBlank()
                        || request.getSubscriptionId().equals(next.getSubscriptionId()))) {
                    observer.onNext(WatchNotificationsResponse.newBuilder().setNotification(next)
                            .build());
                }
            }
            observer.onCompleted();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            observer.onError(Status.CANCELLED.withDescription("watch interrupted").asRuntimeException());
        }
    }

    // ------------------------------------------------------------ subscriptions

    @Override
    public void createSubscription(CreateSubscriptionRequest request,
            StreamObserver<CreateSubscriptionResponse> observer) {
        unary(observer, () -> {
            if (clientState.isBlank()) {
                throw Status.FAILED_PRECONDITION.withDescription("no clientState is configured; set "
                        + NotificationReceiver.ENV_CLIENT_STATE).asRuntimeException();
            }
            return CreateSubscriptionResponse.newBuilder().setSubscription(
                    subscriptions(request.getConnectionId()).create(request, clientState,
                            clock.instant())).build();
        });
    }

    @Override
    public void getSubscription(GetSubscriptionRequest request,
            StreamObserver<GetSubscriptionResponse> observer) {
        unary(observer, () -> GetSubscriptionResponse.newBuilder().setSubscription(
                subscriptions(request.getConnectionId()).get(request.getSubscriptionId())).build());
    }

    @Override
    public void listSubscriptions(ListSubscriptionsRequest request,
            StreamObserver<ListSubscriptionsResponse> observer) {
        unary(observer, () -> ListSubscriptionsResponse.newBuilder().addAllSubscriptions(
                subscriptions(request.getConnectionId()).list()).build());
    }

    @Override
    public void updateSubscription(UpdateSubscriptionRequest request,
            StreamObserver<UpdateSubscriptionResponse> observer) {
        unary(observer, () -> UpdateSubscriptionResponse.newBuilder().setSubscription(
                subscriptions(request.getConnectionId()).update(request, clock.instant())).build());
    }

    @Override
    public void reauthorizeSubscription(ReauthorizeSubscriptionRequest request,
            StreamObserver<ReauthorizeSubscriptionResponse> observer) {
        unary(observer, () -> {
            subscriptions(request.getConnectionId()).reauthorize(request.getSubscriptionId());
            return ReauthorizeSubscriptionResponse.getDefaultInstance();
        });
    }

    @Override
    public void deleteSubscription(DeleteSubscriptionRequest request,
            StreamObserver<DeleteSubscriptionResponse> observer) {
        unary(observer, () -> {
            subscriptions(request.getConnectionId()).delete(request.getSubscriptionId());
            return DeleteSubscriptionResponse.getDefaultInstance();
        });
    }

    private GraphSubscriptions subscriptions(String connectionId) {
        return new GraphSubscriptions(microsoft.sessions().resolve(connectionId).graph());
    }

    // ------------------------------------------------------------ helpers

    @FunctionalInterface
    private interface UnaryBody<T> {
        T run() throws Exception;
    }

    private static <T> void unary(StreamObserver<T> observer, UnaryBody<T> body) {
        T response;
        try {
            response = body.run();
        } catch (Throwable t) {
            observer.onError(GraphErrors.toException(t));
            return;
        }
        observer.onNext(response);
        observer.onCompleted();
    }

    private static Duration duration(com.google.protobuf.Duration value) {
        return Duration.ofSeconds(value.getSeconds(), value.getNanos());
    }

    private static Timestamp timestamp(Instant instant) {
        return Timestamp.newBuilder().setSeconds(instant.getEpochSecond()).setNanos(instant.getNano())
                .build();
    }
}
