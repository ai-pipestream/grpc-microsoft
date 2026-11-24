package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.CreateSubscriptionRequest;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import ai.pipestream.microsoft.v1.RenewalAction;
import ai.pipestream.microsoft.v1.Subscription;
import ai.pipestream.microsoft.v1.SubscriptionRenewal;
import ai.pipestream.microsoft.v1.SubscriptionResource;
import com.google.protobuf.Timestamp;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;

import java.time.Clock;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.ReentrantLock;

/**
 * Keeps the Graph subscriptions this server owns alive: renews each before it
 * expires and on a reauthorizationRequired lifecycle notification, and creates
 * a subscription again when Graph no longer has it.
 */
final class SubscriptionRenewer implements AutoCloseable {

    /** Environment variable that turns renewal off when set to {@code off}. */
    static final String ENV_RENEWAL = "MICROSOFT_SUBSCRIPTION_RENEWAL";
    /** Environment variable for how long before its expiration a subscription is renewed. */
    static final String ENV_RENEW_BEFORE_MINUTES = "MICROSOFT_SUBSCRIPTION_RENEW_BEFORE_MINUTES";
    /** How long before its expiration a subscription is renewed, by default. */
    static final Duration DEFAULT_RENEW_BEFORE = Duration.ofHours(24);
    /** How often the expirations are checked. */
    static final Duration CHECK_INTERVAL = Duration.ofMinutes(5);
    /** A repeated reauthorizationRequired within this time of a renewal is not answered again. */
    static final Duration QUIET_AFTER_RENEWAL = Duration.ofMinutes(10);

    private static final System.Logger LOG = System.getLogger(SubscriptionRenewer.class.getName());

    /** The Graph subscription calls of a connection. */
    @FunctionalInterface
    interface Graphs {
        /**
         * The subscription calls of a catalog connection.
         *
         * @param connectionId the connection; empty = the process default
         * @return its subscription calls
         */
        GraphSubscriptions of(String connectionId);
    }

    /** One kept subscription; guarded by {@link #lock}. */
    private static final class Kept {
        private final String connectionId;
        /** Whether it was created with security webhooks; {@code null} = Graph does not say. */
        private final Boolean securityWebhooks;
        private Subscription subscription;
        private RenewalAction lastAction = RenewalAction.RENEWAL_ACTION_UNSPECIFIED;
        private Instant lastActionAt;
        private String replaced = "";
        private String lastError = "";
        private boolean asked;
        private boolean gone;

        private Kept(String connectionId, Boolean securityWebhooks, Subscription subscription) {
            this.connectionId = connectionId;
            this.securityWebhooks = securityWebhooks;
            this.subscription = subscription;
        }
    }

    private final Graphs graphs;
    private final NotificationBus bus;
    private final String clientState;
    private final Clock clock;
    private final Duration renewBefore;
    private final ReentrantLock lock = new ReentrantLock();
    private final Map<String, Kept> kept = new LinkedHashMap<>();
    private volatile Thread timer;
    private volatile Thread listener;

    /**
     * Creates a renewer that keeps nothing yet.
     *
     * @param graphs the subscription calls of each connection
     * @param bus where a "missed" notification goes after a subscription was created again
     * @param clientState the clientState of this server's subscriptions
     * @param clock time source
     * @param renewBefore how long before its expiration a subscription is renewed
     */
    SubscriptionRenewer(Graphs graphs, NotificationBus bus, String clientState, Clock clock,
            Duration renewBefore) {
        if (clientState == null || clientState.isBlank()) {
            throw new IllegalArgumentException("a clientState is required to renew subscriptions");
        }
        this.graphs = graphs;
        this.bus = bus;
        this.clientState = clientState;
        this.clock = clock;
        this.renewBefore = renewBefore;
    }

    /** Whether the environment leaves renewal on. */
    static boolean enabled(String value) {
        return value == null || !value.trim().equalsIgnoreCase("off");
    }

    /** The renewal margin of an environment value in minutes; the default when unset or unreadable. */
    static Duration renewBefore(String minutes) {
        if (minutes == null || minutes.isBlank()) {
            return DEFAULT_RENEW_BEFORE;
        }
        try {
            long value = Long.parseLong(minutes.trim());
            return value > 0 ? Duration.ofMinutes(value) : DEFAULT_RENEW_BEFORE;
        } catch (NumberFormatException e) {
            return DEFAULT_RENEW_BEFORE;
        }
    }

    // ------------------------------------------------------------ what is kept

    /** Starts keeping a subscription this server just created. */
    void created(String connectionId, boolean securityWebhooks, Subscription subscription) {
        lock.lock();
        try {
            kept.put(subscription.getId(), new Kept(connectionId, securityWebhooks, subscription));
        } finally {
            lock.unlock();
        }
    }

    /**
     * Takes note of a subscription Graph reported: refreshes it when kept, and
     * starts keeping it when it carries this server's clientState.
     *
     * @return whether it is new to the renewer
     */
    boolean seen(String connectionId, GraphSubscriptions.Entry entry) {
        Subscription subscription = entry.subscription();
        lock.lock();
        try {
            Kept known = kept.get(subscription.getId());
            if (known != null) {
                known.subscription = subscription;
                return false;
            }
            if (!clientState.equals(entry.clientState())) {
                return false;
            }
            kept.put(subscription.getId(), new Kept(connectionId, null, subscription));
            return true;
        } finally {
            lock.unlock();
        }
    }

    /** Refreshes a kept subscription after a caller updated it. */
    void updated(Subscription subscription) {
        lock.lock();
        try {
            Kept known = kept.get(subscription.getId());
            if (known != null) {
                known.subscription = subscription;
            }
        } finally {
            lock.unlock();
        }
    }

    /** Stops keeping a subscription. */
    void deleted(String subscriptionId) {
        lock.lock();
        try {
            kept.remove(subscriptionId);
        } finally {
            lock.unlock();
        }
    }

    /**
     * Starts keeping every subscription of a connection that carries this
     * server's clientState.
     *
     * @param connectionId the connection; empty = the process default
     * @return how many were new to the renewer
     * @throws Exception when Graph cannot be asked
     */
    int adopt(String connectionId) throws Exception {
        int adopted = 0;
        for (GraphSubscriptions.Entry entry : graphs.of(connectionId).list()) {
            if (seen(connectionId, entry)) {
                adopted++;
            }
        }
        return adopted;
    }

    /** Every kept subscription, in subscription id order. */
    List<SubscriptionRenewal> renewals() {
        List<SubscriptionRenewal> out = new ArrayList<>();
        lock.lock();
        try {
            for (Kept one : kept.values()) {
                SubscriptionRenewal.Builder renewal = SubscriptionRenewal.newBuilder()
                        .setSubscription(one.subscription).setConnectionId(one.connectionId)
                        .setRenewAt(timestamp(renewAt(one))).setLastAction(one.lastAction)
                        .setReplacedSubscriptionId(one.replaced).setLastError(one.lastError);
                if (one.lastActionAt != null) {
                    renewal.setLastActionTime(timestamp(one.lastActionAt));
                }
                out.add(renewal.build());
            }
        } finally {
            lock.unlock();
        }
        out.sort(Comparator.comparing(renewal -> renewal.getSubscription().getId()));
        return out;
    }

    // ------------------------------------------------------------ keeping

    /** Renews every kept subscription that is due. */
    void runOnce() {
        lock.lock();
        try {
            Instant now = clock.instant();
            for (Kept one : List.copyOf(kept.values())) {
                if (one.gone) {
                    recreate(one, now);
                } else if (one.asked) {
                    renew(one, RenewalAction.RENEWAL_ACTION_REAUTHORIZED, now);
                } else if (!now.isBefore(renewAt(one))) {
                    renew(one, RenewalAction.RENEWAL_ACTION_RENEWED, now);
                }
            }
        } finally {
            lock.unlock();
        }
    }

    /** Answers a lifecycle notification of a kept subscription. */
    void lifecycle(ChangeNotification notification) {
        LifecycleEvent event = notification.getLifecycleEvent();
        if (event != LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED
                && event != LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED) {
            return;
        }
        lock.lock();
        try {
            Kept one = kept.get(notification.getSubscriptionId());
            if (one == null) {
                return;
            }
            Instant now = clock.instant();
            if (event == LifecycleEvent.LIFECYCLE_EVENT_SUBSCRIPTION_REMOVED) {
                recreate(one, now);
            } else if (one.lastActionAt == null || !one.lastError.isEmpty()
                    || !now.isBefore(one.lastActionAt.plus(QUIET_AFTER_RENEWAL))) {
                renew(one, RenewalAction.RENEWAL_ACTION_REAUTHORIZED, now);
            }
        } finally {
            lock.unlock();
        }
    }

    private Instant renewAt(Kept one) {
        if (!one.subscription.hasExpirationDateTime()) {
            return clock.instant();
        }
        Timestamp expires = one.subscription.getExpirationDateTime();
        return Instant.ofEpochSecond(expires.getSeconds(), expires.getNanos()).minus(renewBefore);
    }

    private void renew(Kept one, RenewalAction action, Instant now) {
        String id = one.subscription.getId();
        try {
            one.subscription = graphs.of(one.connectionId).renew(id,
                    one.subscription.getResource(), now);
            done(one, action, now);
            LOG.log(System.Logger.Level.INFO, "renewed subscription {0} ({1})", id, action);
        } catch (Exception e) {
            StatusRuntimeException error = GraphErrors.toException(e);
            if (error.getStatus().getCode() == Status.Code.NOT_FOUND) {
                recreate(one, now);
                return;
            }
            one.asked = action == RenewalAction.RENEWAL_ACTION_REAUTHORIZED;
            failed(one, "renewal", error);
        }
    }

    private void recreate(Kept one, Instant now) {
        Subscription old = one.subscription;
        boolean security = one.securityWebhooks != null ? one.securityWebhooks
                : driveRoot(old.getResource());
        CreateSubscriptionRequest request = CreateSubscriptionRequest.newBuilder()
                .setConnectionId(one.connectionId)
                .setResource(SubscriptionResource.newBuilder().setPath(old.getResource()))
                .addAllChangeTypes(old.getChangeTypesList())
                .setNotificationUrl(old.getNotificationUrl())
                .setLifecycleNotificationUrl(old.getLifecycleNotificationUrl())
                .setIncludeSecurityWebhooks(security).build();
        Subscription created;
        try {
            created = graphs.of(one.connectionId).create(request, clientState, now);
        } catch (Exception e) {
            one.gone = true;
            failed(one, "creating it again", GraphErrors.toException(e));
            return;
        }
        Kept next = new Kept(one.connectionId, security, created);
        next.replaced = old.getId();
        done(next, RenewalAction.RENEWAL_ACTION_RECREATED, now);
        kept.remove(old.getId());
        kept.put(created.getId(), next);
        LOG.log(System.Logger.Level.WARNING, "subscription {0} was gone; created {1} in its place",
                old.getId(), created.getId());
        // Changes between the removal and now were not notified, so watchers run a pass.
        bus.publish(ChangeNotification.newBuilder().setSubscriptionId(created.getId())
                .setResource(created.getResource())
                .setLifecycleEvent(LifecycleEvent.LIFECYCLE_EVENT_MISSED)
                .setSubscriptionExpirationDateTime(created.getExpirationDateTime())
                .setReceivedAt(timestamp(now)).build());
    }

    private static void done(Kept one, RenewalAction action, Instant now) {
        one.lastAction = action;
        one.lastActionAt = now;
        one.lastError = "";
        one.asked = false;
        one.gone = false;
    }

    private static void failed(Kept one, String what, StatusRuntimeException error) {
        one.lastError = error.getStatus().getCode() + ": "
                + java.util.Objects.toString(error.getStatus().getDescription(), "");
        LOG.log(System.Logger.Level.WARNING, "{0} of subscription {1} failed: {2}", what,
                one.subscription.getId(), one.lastError);
    }

    private static boolean driveRoot(String resource) {
        String path = resource.toLowerCase(Locale.ROOT);
        return (path.startsWith("/drives/") || path.startsWith("drives/")) && path.endsWith("/root");
    }

    // ------------------------------------------------------------ threads

    /** Starts answering the lifecycle notifications that arrive on the bus. */
    void answerLifecycles() {
        NotificationBus.Listener lifecycles = bus.listen();
        listener = Thread.ofVirtual().name("subscription-lifecycle").start(() -> {
            try (lifecycles) {
                while (true) {
                    ChangeNotification next = lifecycles.queue().poll(1, TimeUnit.MINUTES);
                    if (next != null) {
                        guarded(() -> lifecycle(next));
                    }
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        });
    }

    /**
     * Starts checking the expirations.
     *
     * @param interval how often
     */
    void checkEvery(Duration interval) {
        timer = Thread.ofVirtual().name("subscription-renewal").start(() -> {
            try {
                while (true) {
                    guarded(this::runOnce);
                    Thread.sleep(interval);
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        });
    }

    private static void guarded(Runnable work) {
        try {
            work.run();
        } catch (RuntimeException e) {
            LOG.log(System.Logger.Level.WARNING, "subscription renewal failed: {0}", e.toString());
        }
    }

    @Override
    public void close() {
        Thread running = timer;
        if (running != null) {
            running.interrupt();
        }
        running = listener;
        if (running != null) {
            running.interrupt();
        }
    }

    private static Timestamp timestamp(Instant instant) {
        return Timestamp.newBuilder().setSeconds(instant.getEpochSecond()).setNanos(instant.getNano())
                .build();
    }
}
