package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.CreateSubscriptionRequest;
import ai.pipestream.microsoft.v1.DeleteSubscriptionRequest;
import ai.pipestream.microsoft.v1.GetSubscriptionRequest;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import ai.pipestream.microsoft.v1.ListSubscriptionRenewalsRequest;
import ai.pipestream.microsoft.v1.ListSubscriptionsRequest;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.PassTrigger;
import ai.pipestream.microsoft.v1.ReauthorizeSubscriptionRequest;
import ai.pipestream.microsoft.v1.RenewalAction;
import ai.pipestream.microsoft.v1.Subscription;
import ai.pipestream.microsoft.v1.SubscriptionChangeType;
import ai.pipestream.microsoft.v1.SubscriptionRenewal;
import ai.pipestream.microsoft.v1.SubscriptionResource;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.UpdateSubscriptionRequest;
import ai.pipestream.microsoft.v1.WatchChangesRequest;
import ai.pipestream.microsoft.v1.WatchChangesResponse;
import ai.pipestream.microsoft.v1.WatchNotificationsRequest;
import ai.pipestream.microsoft.v1.WatchNotificationsResponse;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.google.protobuf.Timestamp;
import io.grpc.Context;
import io.grpc.StatusRuntimeException;
import io.grpc.stub.StreamObserver;

import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;

/**
 * Every {@code MicrosoftChangeService} rpc in one ordered run, shared by the
 * live recording and the replay. Each step's Graph exchanges, and the change
 * notifications Graph posted to the receiver during it, are saved per step.
 */
final class ChangeScenarios {

    /** How long a live step waits for Graph to post a notification. */
    static final Duration NOTIFICATION_WAIT = Duration.ofMinutes(6);
    /** How long a step waits for a WatchChanges pass. */
    static final Duration PASS_WAIT = Duration.ofMinutes(10);
    /** Lifetime of the subscription that should make Graph ask for a reauthorization. */
    static final Duration SHORT_LIFETIME = Duration.ofMinutes(50);
    /** How long after that subscription was created a live run waits for Graph to ask. */
    static final Duration LIFECYCLE_WAIT = Duration.ofMinutes(17);
    /** Renewal margin of the run's renewer: longer than any lifetime, so every check renews. */
    static final Duration RENEW_BEFORE = Duration.ofDays(60);
    /** Step that waits for a lifecycle notification. */
    static final String LIFECYCLE_STEP = "12-lifecycle-renewal";

    /** Where the notifications of a step come from. */
    interface Notifications {
        /**
         * Makes the receiver see this step's notifications: live waits for Graph,
         * replay posts the recorded ones.
         */
        void arrive(String stepId) throws Exception;

        /** Live: waits until no notification has arrived for {@code quiet}; replay: nothing. */
        void settle(Duration quiet) throws Exception;
    }

    /**
     * The test data of a run.
     *
     * @param driveId drive the run watches
     * @param suffix unique run suffix
     * @param startedAt when the live run started; base of every requested expiration
     * @param notificationUrl public URL of the receiver's notification path
     * @param lifecycleUrl public URL of the receiver's lifecycle path
     */
    record Backend(String driveId, String suffix, Instant startedAt, String notificationUrl,
            String lifecycleUrl) {

        /** The clientState of the run. */
        String clientState() {
            return "pipestream-live-change-" + suffix;
        }

        /** Name of the file the run writes. */
        String fileName() {
            return "pipestream-live-change-" + suffix + ".txt";
        }
    }

    /** One step. */
    record Scenario(String id, String rpc, String label, WriteScenarios.Step step) {
    }

    private final MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub blocking;
    private final MicrosoftChangeServiceGrpc.MicrosoftChangeServiceStub async;
    private final GraphClient graph;
    private final Backend backend;
    private final NotificationBus receiverBus;
    private final NotificationBus serviceBus;
    private final Notifications notifications;
    private final SubscriptionRenewer renewer;
    private final Supplier<SubscriptionRenewer> restarted;
    private final BlockingQueue<ChangeNotification> received = new LinkedBlockingQueue<>();
    private final BlockingQueue<ChangeNotification> lifecycles = new LinkedBlockingQueue<>();

    private String subscriptionId;
    private String shortLivedId;
    private long shortLivedCreated;
    private String fileId;
    private Context.CancellableContext watch;
    private final BlockingQueue<WatchChangesResponse> watched = new LinkedBlockingQueue<>();

    ChangeScenarios(MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub blocking,
            MicrosoftChangeServiceGrpc.MicrosoftChangeServiceStub async, GraphClient graph,
            Backend backend, NotificationBus receiverBus, NotificationBus serviceBus,
            Notifications notifications, SubscriptionRenewer renewer,
            Supplier<SubscriptionRenewer> restarted) {
        this.blocking = blocking;
        this.async = async;
        this.graph = graph;
        this.backend = backend;
        this.receiverBus = receiverBus;
        this.serviceBus = serviceBus;
        this.notifications = notifications;
        this.renewer = renewer;
        this.restarted = restarted;
    }

    /**
     * The renewer of a run: its clock stands at the run's start, so the
     * expirations it asks for are the same when the run is replayed.
     */
    static SubscriptionRenewer renewer(SubscriptionRenewer.Graphs graphs, NotificationBus bus,
            Backend backend) {
        return new SubscriptionRenewer(graphs, bus, backend.clientState(),
                java.time.Clock.fixed(backend.startedAt(), java.time.ZoneOffset.UTC), RENEW_BEFORE);
    }

    /** Every step, in run order. */
    List<Scenario> all() {
        return List.of(
                new Scenario("01-create-subscription", "CreateSubscription",
                        "CreateSubscription (drive root, security webhooks)", this::create),
                new Scenario("02-create-short-lived", "CreateSubscription",
                        "CreateSubscription (expires within the hour)", this::createShortLived),
                new Scenario("03-get-subscription", "GetSubscription", "GetSubscription", this::get),
                new Scenario("04-list-subscriptions", "ListSubscriptions", "ListSubscriptions",
                        this::list),
                new Scenario("05-update-subscription", "UpdateSubscription",
                        "UpdateSubscription (renew)", this::update),
                new Scenario("06-reauthorize-subscription", "ReauthorizeSubscription",
                        "ReauthorizeSubscription", this::reauthorize),
                new Scenario("07-watch-first-pass", "WatchChanges", "WatchChanges (first pass)",
                        this::watchStart),
                new Scenario("08-write-file", "", "write a file (Graph helper)", this::writeFile),
                new Scenario("09-notified-pass", "WatchNotifications",
                        "WatchNotifications + WatchChanges pass started by a notification",
                        this::notifiedPass),
                new Scenario("10-permission-notification", "",
                        "notification after a sharing link is added (security webhooks)",
                        this::permissionNotification),
                new Scenario("11-delete-file", "", "delete the file (Graph helper)",
                        this::deleteFile),
                new Scenario(LIFECYCLE_STEP, "",
                        "renewal after Graph asks for a reauthorization (lifecycle notification)",
                        this::lifecycleRenewal),
                new Scenario("13-delete-short-lived", "DeleteSubscription",
                        "DeleteSubscription (the short-lived one)", this::deleteShortLived),
                new Scenario("14-renew-before-expiry", "ListSubscriptionRenewals",
                        "ListSubscriptionRenewals after the server renewed a subscription near its "
                                + "expiration", this::renewBeforeExpiry),
                new Scenario("15-find-after-restart", "",
                        "a restarted server finds its subscription by clientState",
                        this::findAfterRestart),
                new Scenario("16-create-again", "",
                        "the server creates a subscription again that Graph no longer has",
                        this::createAgain),
                new Scenario("17-delete-subscription", "DeleteSubscription", "DeleteSubscription",
                        this::delete));
    }

    /** Runs one step. */
    WriteScenarios.Outcome run(Scenario scenario) {
        try {
            return new WriteScenarios.Outcome(WriteScenarios.Status.PASS, scenario.step().run());
        } catch (Skip e) {
            return new WriteScenarios.Outcome(WriteScenarios.Status.SKIPPED, e.getMessage());
        } catch (StatusRuntimeException e) {
            return new WriteScenarios.Outcome(WriteScenarios.Status.FAIL,
                    e.getStatus().getCode() + " " + e.getStatus().getDescription());
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            return new WriteScenarios.Outcome(WriteScenarios.Status.FAIL, "interrupted");
        } catch (Exception | AssertionError e) {
            return new WriteScenarios.Outcome(WriteScenarios.Status.FAIL,
                    e.getClass().getSimpleName() + ": " + e.getMessage());
        }
    }

    /** Stops any open stream. */
    void close() {
        if (watch != null) {
            watch.cancel(null);
        }
    }

    // ------------------------------------------------------------ steps

    private String create() {
        Subscription created = blocking.createSubscription(CreateSubscriptionRequest.newBuilder()
                .setResource(SubscriptionResource.newBuilder().setDriveRoot(backend.driveId()))
                .addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED)
                .setNotificationUrl(backend.notificationUrl())
                .setLifecycleNotificationUrl(backend.lifecycleUrl())
                .setExpirationDateTime(timestamp(backend.startedAt().plus(Duration.ofDays(1))))
                .setIncludeSecurityWebhooks(true)
                .build()).getSubscription();
        check(!created.getId().isBlank(), "Graph answered no subscription id");
        check(created.getResource().endsWith("/root"), "resource " + created.getResource());
        check(created.getChangeTypesList().contains(
                SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED), "change types "
                + created.getChangeTypesList());
        subscriptionId = created.getId();
        return "subscription created, expires " + created.getExpirationDateTime().getSeconds();
    }

    private String createShortLived() {
        Subscription created = blocking.createSubscription(CreateSubscriptionRequest.newBuilder()
                .setResource(SubscriptionResource.newBuilder().setDriveRoot(backend.driveId()))
                .addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED)
                .setNotificationUrl(backend.notificationUrl())
                .setLifecycleNotificationUrl(backend.lifecycleUrl())
                .setExpirationDateTime(timestamp(backend.startedAt().plus(SHORT_LIFETIME)))
                .build()).getSubscription();
        check(!created.getId().isBlank(), "Graph answered no subscription id");
        shortLivedId = created.getId();
        shortLivedCreated = System.nanoTime();
        return "subscription created, expires " + SHORT_LIFETIME.toMinutes()
                + " min after the run started";
    }

    private String get() {
        Subscription got = blocking.getSubscription(GetSubscriptionRequest.newBuilder()
                .setSubscriptionId(need(subscriptionId, "a subscription")).build()).getSubscription();
        check(got.getId().equals(subscriptionId), "GetSubscription answered another id");
        return "read back";
    }

    private String list() {
        List<Subscription> all = blocking.listSubscriptions(ListSubscriptionsRequest.getDefaultInstance())
                .getSubscriptionsList();
        String id = need(subscriptionId, "a subscription");
        check(all.stream().anyMatch(sub -> sub.getId().equals(id)),
                "ListSubscriptions did not include the new subscription");
        return all.size() + " subscription(s) listed";
    }

    private String update() {
        Timestamp wanted = timestamp(backend.startedAt().plus(Duration.ofDays(2)));
        Subscription updated = blocking.updateSubscription(UpdateSubscriptionRequest.newBuilder()
                .setSubscriptionId(need(subscriptionId, "a subscription"))
                .setExpirationDateTime(wanted).build()).getSubscription();
        check(updated.getExpirationDateTime().getSeconds() == wanted.getSeconds(),
                "expiration is " + updated.getExpirationDateTime().getSeconds() + ", wanted "
                        + wanted.getSeconds());
        return "renewed by one day";
    }

    private String reauthorize() {
        blocking.reauthorizeSubscription(ReauthorizeSubscriptionRequest.newBuilder()
                .setSubscriptionId(need(subscriptionId, "a subscription")).build());
        return "reauthorized";
    }

    private String watchStart() throws Exception {
        watch = Context.current().withCancellation();
        WatchChangesRequest request = WatchChangesRequest.newBuilder()
                .setSync(SyncRequest.newBuilder().addDriveIds(backend.driveId()))
                .setPollInterval(com.google.protobuf.Duration.newBuilder().setSeconds(3600))
                .setMinInterval(com.google.protobuf.Duration.getDefaultInstance())
                .build();
        watch.run(() -> async.watchChanges(request, new StreamObserver<>() {
            @Override
            public void onNext(WatchChangesResponse value) {
                watched.add(value);
            }

            @Override
            public void onError(Throwable t) {
                // A cancelled watch ends with CANCELLED.
            }

            @Override
            public void onCompleted() {
                // Not expected while the watch is open.
            }
        }));
        List<WatchChangesResponse> pass = awaitPass(PassTrigger.PASS_TRIGGER_START);
        long sequence = 0;
        for (WatchChangesResponse message : pass) {
            check(message.getSequence() == sequence + 1, "sequence jumped from " + sequence
                    + " to " + message.getSequence());
            sequence = message.getSequence();
        }
        WatchChangesResponse last = pass.get(pass.size() - 1);
        check(!last.getResumeCursor().isBlank(), "the first pass ended without a resume cursor");
        return "first pass: " + last.getPassCompleted().getChangeCount() + " change(s), "
                + pass.size() + " message(s)";
    }

    private String writeFile() throws Exception {
        need(watch, "an open watch");
        byte[] content = ChangeRecordings.content(backend.suffix());
        JsonNode item = graph.putBytes("/drives/" + backend.driveId() + "/root:/"
                + backend.fileName() + ":/content", content, "text/plain");
        fileId = item.path("id").asText("");
        check(!fileId.isBlank(), "the upload answered no item id");
        return "uploaded " + content.length + " bytes";
    }

    private String notifiedPass() throws Exception {
        need(fileId, "the written file");
        Context.CancellableContext stream = Context.current().withCancellation();
        BlockingQueue<WatchNotificationsResponse> seen = new LinkedBlockingQueue<>();
        try {
            stream.run(() -> async.watchNotifications(WatchNotificationsRequest.newBuilder()
                    .setSubscriptionId(subscriptionId).build(), collector(seen)));
            notifications.arrive("09-notified-pass");
            ChangeNotification first = awaitNotification(NOTIFICATION_WAIT);
            Thread.sleep(200);
            serviceBus.publish(first);
            WatchNotificationsResponse echoed = seen.poll(30, TimeUnit.SECONDS);
            check(echoed != null, "WatchNotifications did not stream the notification");
            check(echoed.getNotification().getSubscriptionId().equals(subscriptionId),
                    "WatchNotifications streamed another subscription's notification");
            List<WatchChangesResponse> pass = awaitPass(PassTrigger.PASS_TRIGGER_NOTIFICATION);
            MicrosoftChange change = pass.stream().filter(WatchChangesResponse::hasChange)
                    .map(WatchChangesResponse::getChange)
                    .filter(c -> c.getEntity().getDriveItem().getName().equals(backend.fileName()))
                    .findFirst().orElse(null);
            check(change != null, "the notified pass did not emit the written file");
            check(change.getSource() == ChangeSource.CHANGE_SOURCE_NOTIFICATION,
                    "the file's change source is " + change.getSource());
            return "notification " + first.getChangeType() + " started a pass with "
                    + pass.get(pass.size() - 1).getPassCompleted().getChangeCount() + " change(s)";
        } finally {
            stream.cancel(null);
            close();
        }
    }

    private String permissionNotification() throws Exception {
        String id = need(fileId, "the written file");
        notifications.settle(Duration.ofSeconds(90));
        received.clear();
        ObjectNode body = Recordings.JSON.createObjectNode().put("type", "view")
                .put("scope", "organization");
        graph.post("/drives/" + backend.driveId() + "/items/" + id + "/createLink", body);
        notifications.arrive("10-permission-notification");
        ChangeNotification next = awaitNotification(NOTIFICATION_WAIT);
        return "notification " + next.getChangeType() + " after the sharing link was added";
    }

    private String deleteFile() throws Exception {
        graph.delete("/drives/" + backend.driveId() + "/items/" + need(fileId, "the written file"));
        fileId = null;
        return "deleted";
    }

    private String lifecycleRenewal() throws Exception {
        String id = need(shortLivedId, "the short-lived subscription");
        notifications.arrive(LIFECYCLE_STEP);
        long deadline = shortLivedCreated + LIFECYCLE_WAIT.toNanos();
        ChangeNotification asked = null;
        while (asked == null) {
            ChangeNotification next = lifecycles.poll(500, TimeUnit.MILLISECONDS);
            if (next != null && next.getSubscriptionId().equals(id)) {
                asked = next;
            } else if (System.nanoTime() > deadline) {
                throw new Skip("Graph sent no lifecycle notification within "
                        + LIFECYCLE_WAIT.toMinutes() + " min for a subscription that expires in "
                        + SHORT_LIFETIME.toMinutes() + " min");
            }
        }
        check(asked.getLifecycleEvent() == LifecycleEvent.LIFECYCLE_EVENT_REAUTHORIZATION_REQUIRED,
                "the lifecycle notification is " + asked.getLifecycleEvent());
        long waited = TimeUnit.NANOSECONDS.toSeconds(System.nanoTime() - shortLivedCreated);
        serviceBus.publish(asked);
        SubscriptionRenewal renewal = null;
        long until = System.nanoTime() + TimeUnit.SECONDS.toNanos(60);
        while (System.nanoTime() < until) {
            renewal = renewal(id);
            if (renewal != null
                    && renewal.getLastAction() != RenewalAction.RENEWAL_ACTION_UNSPECIFIED) {
                break;
            }
            Thread.sleep(100);
        }
        check(renewal != null, "the server no longer keeps the short-lived subscription");
        check(renewal.getLastAction() == RenewalAction.RENEWAL_ACTION_REAUTHORIZED,
                "the server answered with " + renewal.getLastAction() + " " + renewal.getLastError());
        check(renewal.getSubscription().getExpirationDateTime().getSeconds() == longest(renewal),
                "the renewed expiration is " + renewal.getSubscription().getExpirationDateTime()
                        .getSeconds() + ", wanted " + longest(renewal));
        return "Graph asked for a reauthorization (waited " + (waited < 60 ? "under a minute"
                : waited / 60 + " min") + "); the server renewed the subscription";
    }

    private String deleteShortLived() {
        blocking.deleteSubscription(DeleteSubscriptionRequest.newBuilder()
                .setSubscriptionId(need(shortLivedId, "the short-lived subscription")).build());
        check(renewal(shortLivedId) == null, "the server still keeps the deleted subscription");
        shortLivedId = null;
        return "deleted";
    }

    private String renewBeforeExpiry() {
        String id = need(subscriptionId, "a subscription");
        SubscriptionRenewal before = renewal(id);
        check(before != null, "the server does not keep the subscription it created");
        check(before.getLastAction() == RenewalAction.RENEWAL_ACTION_UNSPECIFIED,
                "the server already acted on it: " + before.getLastAction());
        renewer.runOnce();
        SubscriptionRenewal after = renewal(id);
        check(after != null && after.getLastAction() == RenewalAction.RENEWAL_ACTION_RENEWED,
                "the server did not renew it: " + (after == null ? "not kept"
                        : after.getLastAction() + " " + after.getLastError()));
        long wanted = longest(after);
        check(after.getSubscription().getExpirationDateTime().getSeconds() == wanted,
                "the renewed expiration is " + after.getSubscription().getExpirationDateTime()
                        .getSeconds() + ", wanted " + wanted);
        Subscription got = blocking.getSubscription(GetSubscriptionRequest.newBuilder()
                .setSubscriptionId(id).build()).getSubscription();
        check(got.getExpirationDateTime().getSeconds() == wanted,
                "Graph reports the expiration " + got.getExpirationDateTime().getSeconds()
                        + ", wanted " + wanted);
        return "renewed to the longest expiration, "
                + GraphSubscriptions.maxLifetime(got.getResource()).toDays() + " days";
    }

    private String findAfterRestart() throws Exception {
        String id = need(subscriptionId, "a subscription");
        SubscriptionRenewer fresh = restarted.get();
        int found = fresh.adopt("");
        check(fresh.renewals().stream().anyMatch(one -> one.getSubscription().getId().equals(id)),
                "a new renewer did not find the subscription");
        return found + " subscription(s) with this run's clientState found";
    }

    private String createAgain() throws Exception {
        String old = need(subscriptionId, "a subscription");
        graph.delete("/subscriptions/" + old);
        try (NotificationBus.Listener bus = serviceBus.listen()) {
            renewer.runOnce();
            SubscriptionRenewal renewal = renewals().stream()
                    .filter(one -> one.getReplacedSubscriptionId().equals(old)).findFirst()
                    .orElse(null);
            SubscriptionRenewal stale = renewal(old);
            check(renewal != null, "the server did not create it again"
                    + (stale == null ? "" : ": " + stale.getLastError()));
            subscriptionId = renewal.getSubscription().getId();
            check(renewal.getLastAction() == RenewalAction.RENEWAL_ACTION_RECREATED,
                    "the last action is " + renewal.getLastAction());
            check(!subscriptionId.equals(old), "the new subscription has the old id");
            check(stale == null, "the server still keeps the removed subscription");
            ChangeNotification missed = bus.queue().poll(5, TimeUnit.SECONDS);
            check(missed != null && missed.getSubscriptionId().equals(subscriptionId)
                            && missed.getLifecycleEvent() == LifecycleEvent.LIFECYCLE_EVENT_MISSED,
                    "watchers were not told to run a pass");
        }
        return "Graph answered NOT_FOUND to the renewal; the server created the subscription again";
    }

    private String delete() {
        blocking.deleteSubscription(DeleteSubscriptionRequest.newBuilder()
                .setSubscriptionId(need(subscriptionId, "a subscription")).build());
        subscriptionId = null;
        check(renewals().isEmpty(), "the server still keeps " + renewals().size()
                + " subscription(s)");
        return "deleted";
    }

    private List<SubscriptionRenewal> renewals() {
        return blocking.listSubscriptionRenewals(ListSubscriptionRenewalsRequest.getDefaultInstance())
                .getRenewalsList();
    }

    private SubscriptionRenewal renewal(String id) {
        return renewals().stream().filter(one -> one.getSubscription().getId().equals(id))
                .findFirst().orElse(null);
    }

    /** The expiration the run's renewer asks for, in epoch seconds. */
    private long longest(SubscriptionRenewal renewal) {
        return backend.startedAt().plus(GraphSubscriptions.maxLifetime(
                renewal.getSubscription().getResource())).getEpochSecond();
    }

    /** Removes what a failed run left behind. */
    void cleanUp() {
        close();
        try {
            if (fileId != null) {
                graph.delete("/drives/" + backend.driveId() + "/items/" + fileId);
            }
        } catch (Exception e) {
            System.out.println("[ChangeScenarios] file clean-up failed: " + e.getMessage());
        }
        for (String id : java.util.Arrays.asList(subscriptionId, shortLivedId)) {
            try {
                if (id != null) {
                    blocking.deleteSubscription(DeleteSubscriptionRequest.newBuilder()
                            .setSubscriptionId(id).build());
                }
            } catch (RuntimeException e) {
                System.out.println("[ChangeScenarios] subscription clean-up failed: "
                        + e.getMessage());
            }
        }
    }

    // ------------------------------------------------------------ helpers

    /** Opens the receiver-side listener; call before the first step. */
    NotificationBus.Listener listen() {
        NotificationBus.Listener listener = receiverBus.listen();
        Thread.ofVirtual().start(() -> {
            try {
                while (true) {
                    ChangeNotification next = listener.queue().take();
                    (next.getLifecycleEvent() == LifecycleEvent.LIFECYCLE_EVENT_UNSPECIFIED
                            ? received : lifecycles).add(next);
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
            }
        });
        return listener;
    }

    private ChangeNotification awaitNotification(Duration wait) throws InterruptedException {
        long deadline = System.nanoTime() + wait.toNanos();
        while (System.nanoTime() < deadline) {
            ChangeNotification next = received.poll(500, TimeUnit.MILLISECONDS);
            if (next != null && next.getSubscriptionId().equals(subscriptionId)) {
                return next;
            }
        }
        throw new AssertionError("no change notification within " + wait.toSeconds() + " s");
    }

    private List<WatchChangesResponse> awaitPass(PassTrigger trigger) throws InterruptedException {
        List<WatchChangesResponse> pass = new ArrayList<>();
        long deadline = System.nanoTime() + PASS_WAIT.toNanos();
        while (System.nanoTime() < deadline) {
            WatchChangesResponse next = watched.poll(500, TimeUnit.MILLISECONDS);
            if (next == null) {
                continue;
            }
            pass.add(next);
            if (next.hasPassCompleted()) {
                check(next.getPassCompleted().getTrigger() == trigger, "pass started by "
                        + next.getPassCompleted().getTrigger() + ", expected " + trigger);
                return pass;
            }
        }
        throw new AssertionError("no " + trigger + " pass within " + PASS_WAIT.toMinutes() + " min");
    }

    private static <T> StreamObserver<T> collector(BlockingQueue<T> into) {
        return new StreamObserver<>() {
            @Override
            public void onNext(T value) {
                into.add(value);
            }

            @Override
            public void onError(Throwable t) {
                // A cancelled stream ends with CANCELLED.
            }

            @Override
            public void onCompleted() {
                // Nothing to do.
            }
        };
    }

    private static Timestamp timestamp(Instant instant) {
        return Timestamp.newBuilder().setSeconds(instant.getEpochSecond()).build();
    }

    private static <T> T need(T value, String what) {
        if (value == null) {
            throw new Skip("needs " + what + ", which an earlier step did not produce");
        }
        return value;
    }

    private static void check(boolean condition, String message) {
        if (!condition) {
            throw new AssertionError(message);
        }
    }

    private static final class Skip extends RuntimeException {
        private static final long serialVersionUID = 1L;

        Skip(String message) {
            super(message);
        }
    }
}
