package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.CreateSubscriptionRequest;
import ai.pipestream.microsoft.v1.DeleteSubscriptionRequest;
import ai.pipestream.microsoft.v1.DirectoryKind;
import ai.pipestream.microsoft.v1.DirectorySnapshot;
import ai.pipestream.microsoft.v1.DirectoryWatch;
import ai.pipestream.microsoft.v1.GetSubscriptionRequest;
import ai.pipestream.microsoft.v1.GroupChange;
import ai.pipestream.microsoft.v1.LifecycleEvent;
import ai.pipestream.microsoft.v1.ListSubscriptionRenewalsRequest;
import ai.pipestream.microsoft.v1.ListSubscriptionsRequest;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftChangeServiceGrpc;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.PassTrigger;
import ai.pipestream.microsoft.v1.ReauthorizeSubscriptionRequest;
import ai.pipestream.microsoft.v1.RenewalAction;
import ai.pipestream.microsoft.v1.SiteChange;
import ai.pipestream.microsoft.v1.SiteChangeObject;
import ai.pipestream.microsoft.v1.SiteChangeScope;
import ai.pipestream.microsoft.v1.SiteChangeSnapshot;
import ai.pipestream.microsoft.v1.SiteChangeType;
import ai.pipestream.microsoft.v1.SiteChangeWatch;
import ai.pipestream.microsoft.v1.Subscription;
import ai.pipestream.microsoft.v1.SubscriptionChangeType;
import ai.pipestream.microsoft.v1.SubscriptionRenewal;
import ai.pipestream.microsoft.v1.SubscriptionResource;
import ai.pipestream.microsoft.v1.SyncCursor;
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
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.stub.StreamObserver;

import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
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
    static final Duration LIFECYCLE_WAIT = Duration.ofMinutes(2);
    /** The longest such wait: the subscription must still exist when the wait ends. */
    static final Duration LONGEST_LIFECYCLE_WAIT = SHORT_LIFETIME.minusMinutes(4);
    /** Renewal margin of the run's renewer: longer than any lifetime, so every check renews. */
    static final Duration RENEW_BEFORE = Duration.ofDays(60);
    /** Step that waits for a lifecycle notification. */
    static final String LIFECYCLE_STEP = "12-lifecycle-renewal";
    /** Step that waits for the notification of a change to an Entra group. */
    static final String DIRECTORY_STEP = "24-directory-notification";
    /** How long a live step waits for Graph to post a notification about the directory. */
    static final Duration DIRECTORY_NOTIFICATION_WAIT = Duration.ofMinutes(15);
    /** Tries a step gives the directory to show what the run just wrote. */
    static final int DIRECTORY_TRIES = 30;

    /** Where the notifications of a step come from. */
    interface Notifications {
        /**
         * Makes the receiver see this step's notifications: live waits for Graph,
         * replay posts the recorded ones.
         */
        void arrive(String stepId) throws Exception;

        /** Live: waits until no notification has arrived for {@code quiet}; replay: nothing. */
        void settle(Duration quiet) throws Exception;

        /** Live: waits {@code time} for Microsoft to catch up with a write; replay: nothing. */
        void pause(Duration time) throws Exception;
    }

    /**
     * The test data of a run.
     *
     * @param driveId drive the run watches
     * @param suffix unique run suffix
     * @param startedAt when the live run started; base of every requested expiration
     * @param notificationUrl public URL of the receiver's notification path
     * @param lifecycleUrl public URL of the receiver's lifecycle path
     * @param siteId the site of the drive; empty when the run does not know it
     */
    record Backend(String driveId, String suffix, Instant startedAt, String notificationUrl,
            String lifecycleUrl, String siteId) {

        /** The clientState of the run. */
        String clientState() {
            return "pipestream-live-change-" + suffix;
        }

        /** Name of the file the run writes. */
        String fileName() {
            return "pipestream-live-change-" + suffix + ".txt";
        }

        /** Name of the Entra group the run makes. */
        String groupName() {
            return "pipestream-live-change-" + suffix + "-entra";
        }
    }

    /** One step. */
    record Scenario(String id, String rpc, String label, WriteScenarios.Step step) {
    }

    private final MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub blocking;
    private final MicrosoftChangeServiceGrpc.MicrosoftChangeServiceStub async;
    private final MicrosoftServiceGrpc.MicrosoftServiceBlockingStub microsoft;
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
    private Duration lifecycleWait = LIFECYCLE_WAIT;
    private String directoryCursor;
    private String userPrincipalName;
    private String fileId;
    private String entraGroupId;
    private String directorySubscriptionId;
    private Context.CancellableContext watch;
    private final BlockingQueue<WatchChangesResponse> watched = new LinkedBlockingQueue<>();

    ChangeScenarios(MicrosoftChangeServiceGrpc.MicrosoftChangeServiceBlockingStub blocking,
            MicrosoftChangeServiceGrpc.MicrosoftChangeServiceStub async,
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub microsoft, GraphClient graph,
            Backend backend, NotificationBus receiverBus, NotificationBus serviceBus,
            Notifications notifications, SubscriptionRenewer renewer,
            Supplier<SubscriptionRenewer> restarted) {
        this.blocking = blocking;
        this.async = async;
        this.microsoft = microsoft;
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

    /**
     * Sets how long the run waits for Graph to ask for a reauthorization.
     *
     * @param wait the wait, cut to {@link #LONGEST_LIFECYCLE_WAIT}
     */
    void lifecycleWait(Duration wait) {
        lifecycleWait = wait.compareTo(LONGEST_LIFECYCLE_WAIT) > 0 ? LONGEST_LIFECYCLE_WAIT : wait;
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
                        this::delete),
                new Scenario("18-watch-directory", "WatchChanges",
                        "WatchChanges (every Entra user and group, with members)",
                        this::watchDirectory),
                new Scenario("19-watch-directory-resumed", "WatchChanges",
                        "WatchChanges (Entra changes since the cursor)", this::watchDirectoryResumed),
                new Scenario("20-subscribe-user", "CreateSubscription",
                        "CreateSubscription (one Entra user)", this::subscribeToAUser),
                new Scenario("21-subscribe-group", "CreateSubscription",
                        "CreateSubscription (one Entra group)", this::subscribeToAGroup),
                new Scenario("22-subscribe-group-members", "CreateSubscription",
                        "CreateSubscription (the members of one Entra group)",
                        this::subscribeToGroupMembers),
                new Scenario("23-watch-site-log", "WatchChanges",
                        "WatchChanges (the site collection's change log: a sharing link, a group member)",
                        this::watchSiteLog),
                new Scenario(DIRECTORY_STEP, "",
                        "notification after an Entra group gets a member, and the pass that reads it",
                        this::directoryNotification));
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
        long deadline = shortLivedCreated + lifecycleWait.toNanos();
        ChangeNotification asked = null;
        while (asked == null) {
            ChangeNotification next = lifecycles.poll(500, TimeUnit.MILLISECONDS);
            if (next != null && next.getSubscriptionId().equals(id)) {
                asked = next;
            } else if (System.nanoTime() > deadline) {
                throw new Skip("Graph sent no lifecycle notification within "
                        + lifecycleWait.toMinutes() + " min for a subscription that expires in "
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

    private String watchDirectory() throws Exception {
        List<WatchChangesResponse> pass = directoryPass("");
        long users = pass.stream().filter(WatchChangesResponse::hasUserChange).count();
        long groups = pass.stream().filter(WatchChangesResponse::hasGroupChange).count();
        long members = pass.stream().filter(WatchChangesResponse::hasGroupChange)
                .mapToLong(message -> message.getGroupChange().getMemberChangesCount()).sum();
        check(users > 0, "the pass listed no user");
        check(pass.stream().filter(WatchChangesResponse::hasUserChange)
                        .noneMatch(message -> message.getUserChange().getUser().getId().isBlank()),
                "a user change has no id");
        check(pass.stream().filter(WatchChangesResponse::hasGroupChange)
                        .noneMatch(message -> message.getGroupChange().getGroup().getId().isBlank()),
                "a group change has no id");
        checkSnapshot(pass, DirectoryKind.DIRECTORY_KIND_USERS, ChangeSource.CHANGE_SOURCE_CRAWL, users);
        checkSnapshot(pass, DirectoryKind.DIRECTORY_KIND_GROUPS, ChangeSource.CHANGE_SOURCE_CRAWL,
                groups);
        WatchChangesResponse last = pass.get(pass.size() - 1);
        SyncCursor cursor = SyncCursor.parseFrom(Base64.getUrlDecoder().decode(last.getResumeCursor()));
        check(!cursor.getUsersDeltaLink().isBlank() && !cursor.getGroupsDeltaLink().isBlank(),
                "the resume cursor holds no users or groups delta link");
        check(last.getPassCompleted().getChangeCount() >= users + groups,
                "the pass counted " + last.getPassCompleted().getChangeCount() + " change(s)");
        directoryCursor = last.getResumeCursor();
        userPrincipalName = pass.stream().filter(WatchChangesResponse::hasUserChange)
                .map(message -> message.getUserChange().getUser().getUserPrincipalName())
                .filter(name -> !name.isBlank()).findFirst().orElse(null);
        return users + " user(s), " + groups + " group(s) with " + members + " member change(s)";
    }

    private String watchDirectoryResumed() throws Exception {
        List<WatchChangesResponse> pass = directoryPass(need(directoryCursor,
                "the cursor of the directory pass"));
        long users = pass.stream().filter(WatchChangesResponse::hasUserChange).count();
        long groups = pass.stream().filter(WatchChangesResponse::hasGroupChange).count();
        checkSnapshot(pass, DirectoryKind.DIRECTORY_KIND_USERS, ChangeSource.CHANGE_SOURCE_DELTA, users);
        checkSnapshot(pass, DirectoryKind.DIRECTORY_KIND_GROUPS, ChangeSource.CHANGE_SOURCE_DELTA,
                groups);
        return "since the cursor: " + users + " user change(s), " + groups + " group change(s)";
    }

    private String watchSiteLog() throws Exception {
        if (backend.siteId().isBlank()) {
            throw new Skip("the run does not know the site of the drive");
        }
        need(userPrincipalName, "a directory user with a sign-in name");
        WatchChangesRequest.Builder request = WatchChangesRequest.newBuilder()
                .setSiteChanges(SiteChangeWatch.newBuilder().addSiteIds(backend.siteId())
                        .setScope(SiteChangeScope.SITE_CHANGE_SCOPE_SITE_COLLECTION));
        List<WatchChangesResponse> first;
        try {
            first = firstPass(request.build());
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == Status.Code.FAILED_PRECONDITION) {
                throw new Skip("no certificate credential: " + e.getStatus().getDescription());
            }
            throw e;
        }
        SiteChangeSnapshot start = siteSnapshot(first);
        check(start.getSource() == ChangeSource.CHANGE_SOURCE_CRAWL && start.getChangeCount() == 0
                        && first.stream().noneMatch(WatchChangesResponse::hasSiteChange),
                "the first pass of the site did not start at the end of its log");
        String cursor = first.get(first.size() - 1).getResumeCursor();
        check(!SyncCursor.parseFrom(Base64.getUrlDecoder().decode(cursor))
                .getSiteChangeTokensOrDefault(backend.siteId(), "").isBlank(),
                "the resume cursor holds no change token of the site");

        JsonNode item = graph.putBytes("/drives/" + backend.driveId() + "/root:/site-log-"
                + backend.fileName() + ":/content", ChangeRecordings.content(backend.suffix()),
                "text/plain");
        String id = item.path("id").asText("");
        check(!id.isBlank(), "the upload answered no item id");
        try {
            graph.post("/drives/" + backend.driveId() + "/items/" + id + "/createLink",
                    Recordings.JSON.createObjectNode().put("type", "view").put("scope", "organization"));
        } finally {
            graph.delete("/drives/" + backend.driveId() + "/items/" + id);
        }

        int groupId = microsoft.createSiteGroup(ai.pipestream.microsoft.v1.CreateSiteGroupRequest.newBuilder()
                .setSiteId(backend.siteId()).setTitle("pipestream-live-change-" + backend.suffix() + "-group")
                .setDescription("Created by a live test run; safe to delete.").build()).getGroup().getId();
        try {
            int memberId = microsoft.addSiteGroupMember(ai.pipestream.microsoft.v1.AddSiteGroupMemberRequest.newBuilder()
                    .setSiteId(backend.siteId()).setGroupId(groupId)
                    .setUserPrincipalName(userPrincipalName).build()).getMember().getId();
            microsoft.removeSiteGroupMember(ai.pipestream.microsoft.v1.RemoveSiteGroupMemberRequest.newBuilder()
                    .setSiteId(backend.siteId()).setGroupId(groupId).setMemberId(memberId).build());
        } finally {
            microsoft.deleteSiteGroup(ai.pipestream.microsoft.v1.DeleteSiteGroupRequest.newBuilder()
                    .setSiteId(backend.siteId()).setGroupId(groupId).build());
        }

        List<WatchChangesResponse> next = firstPass(request.setResumeCursor(cursor).build());
        SiteChangeSnapshot read = siteSnapshot(next);
        List<SiteChange> changes = next.stream().filter(WatchChangesResponse::hasSiteChange)
                .map(WatchChangesResponse::getSiteChange).toList();
        check(read.getSource() == ChangeSource.CHANGE_SOURCE_DELTA
                && read.getChangeCount() == changes.size(), "the snapshot is " + read.getSource()
                + " with " + read.getChangeCount() + " for " + changes.size() + " change(s)");
        java.util.Map<String, Integer> kinds = new java.util.TreeMap<>();
        changes.forEach(change -> kinds.merge(change.getObject().name().substring(
                "SITE_CHANGE_OBJECT_".length()) + " " + change.getType().name().substring(
                "SITE_CHANGE_TYPE_".length()), 1, Integer::sum));
        check(changes.stream().anyMatch(change -> change.getObject()
                == SiteChangeObject.SITE_CHANGE_OBJECT_ITEM && change.getSecurable().hasListItem()),
                "the log holds no list item change: " + kinds);
        for (SiteChangeType membership : List.of(SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_ADD,
                SiteChangeType.SITE_CHANGE_TYPE_GROUP_MEMBER_DELETE)) {
            check(changes.stream().anyMatch(change -> change.getType() == membership
                    && change.getObject() == SiteChangeObject.SITE_CHANGE_OBJECT_GROUP
                    && change.getGroupId() == groupId),
                    "the log holds no " + membership + " of the run's group: " + kinds);
        }
        return "first pass took the end of the log; after a file was shared and a group got and lost "
                + "a member: " + kinds;
    }

    private static SiteChangeSnapshot siteSnapshot(List<WatchChangesResponse> pass) {
        SiteChangeSnapshot snapshot = pass.stream()
                .filter(WatchChangesResponse::hasSiteChangeSnapshot)
                .map(WatchChangesResponse::getSiteChangeSnapshot).findFirst().orElse(null);
        check(snapshot != null, "the pass sent no snapshot of the site's change log");
        return snapshot;
    }

    /** Opens a watch with the directory on, collects its first pass and closes it. */
    private List<WatchChangesResponse> directoryPass(String resumeCursor) throws Exception {
        return firstPass(WatchChangesRequest.newBuilder()
                .setResumeCursor(resumeCursor)
                .setDirectory(DirectoryWatch.newBuilder().setUsers(true).setGroups(true)
                        .setGroupMembers(true))
                .build());
    }

    /** Opens a watch of the run's drive, collects its first pass and closes it. */
    private List<WatchChangesResponse> firstPass(WatchChangesRequest watched) throws Exception {
        Context.CancellableContext stream = Context.current().withCancellation();
        BlockingQueue<WatchChangesResponse> seen = new LinkedBlockingQueue<>();
        AtomicReference<Throwable> failure = new AtomicReference<>();
        WatchChangesRequest request = watched.toBuilder()
                .setSync(SyncRequest.newBuilder().addDriveIds(backend.driveId()))
                .setPollInterval(com.google.protobuf.Duration.newBuilder().setSeconds(3600))
                .build();
        try {
            stream.run(() -> async.watchChanges(request, new StreamObserver<>() {
                @Override
                public void onNext(WatchChangesResponse value) {
                    seen.add(value);
                }

                @Override
                public void onError(Throwable t) {
                    failure.set(t);
                }

                @Override
                public void onCompleted() {
                    // Not expected while the watch is open.
                }
            }));
            List<WatchChangesResponse> pass = new ArrayList<>();
            long deadline = System.nanoTime() + PASS_WAIT.toNanos();
            while (System.nanoTime() < deadline) {
                WatchChangesResponse next = seen.poll(200, TimeUnit.MILLISECONDS);
                if (next == null) {
                    if (failure.get() != null) {
                        throw Status.fromThrowable(failure.get()).asRuntimeException();
                    }
                    continue;
                }
                pass.add(next);
                if (next.hasPassCompleted()) {
                    return pass;
                }
            }
            throw new AssertionError("no pass within " + PASS_WAIT.toMinutes() + " min");
        } finally {
            stream.cancel(null);
        }
    }

    private static void checkSnapshot(List<WatchChangesResponse> pass, DirectoryKind kind,
            ChangeSource source, long changes) {
        DirectorySnapshot snapshot = pass.stream()
                .filter(WatchChangesResponse::hasDirectorySnapshot)
                .map(WatchChangesResponse::getDirectorySnapshot)
                .filter(one -> one.getKind() == kind).findFirst().orElse(null);
        check(snapshot != null, "the pass sent no snapshot of " + kind);
        check(snapshot.getSource() == source, "the snapshot of " + kind + " is " + snapshot.getSource());
        check(snapshot.getChangeCount() == changes, "the snapshot of " + kind + " counts "
                + snapshot.getChangeCount() + ", the pass held " + changes);
    }

    private String directoryNotification() throws Exception {
        String cursor = need(directoryCursor, "the cursor of the directory pass");
        String userId = graph.get("/users?$top=1&$select=id").path("value").path(0).path("id").asText("");
        check(!userId.isBlank(), "the tenant lists no user");
        String name = backend.groupName();
        try {
            entraGroupId = microsoft.createGroup(ai.pipestream.microsoft.v1.CreateGroupRequest.newBuilder()
                    .setGroup(ai.pipestream.microsoft.v1.GraphGroup.newBuilder().setDisplayName(name)
                            .setMailNickname(name)
                            .setDescription(name + " created by a live test run; safe to delete"))
                    .build()).getGroup().getId();
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == Status.Code.PERMISSION_DENIED) {
                throw new Skip("the app may not write to the directory (Group.ReadWrite.All): "
                        + e.getStatus().getDescription());
            }
            throw e;
        }
        check(!entraGroupId.isBlank(), "Graph answered no group id");
        try {
            int subscribed = subscribeToTheNewGroup();
            notifications.settle(Duration.ofSeconds(30));
            received.clear();
            addMember(userId);
            long asked = System.nanoTime();
            notifications.arrive(DIRECTORY_STEP);
            ChangeNotification notification = awaitNotification(DIRECTORY_NOTIFICATION_WAIT,
                    directorySubscriptionId);
            long waited = Duration.ofNanos(System.nanoTime() - asked).toSeconds();
            check(notification.getChangeType() == SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED,
                    "the notification says " + notification.getChangeType());
            check(notification.getResource().toLowerCase(java.util.Locale.ROOT)
                            .contains(entraGroupId.toLowerCase(java.util.Locale.ROOT)),
                    "the notification names another resource");
            int passes = 0;
            GroupChange seen = null;
            while (seen == null && passes < DIRECTORY_TRIES) {
                passes++;
                seen = directoryPass(cursor).stream().filter(WatchChangesResponse::hasGroupChange)
                        .map(WatchChangesResponse::getGroupChange)
                        .filter(change -> change.getGroup().getId().equals(entraGroupId)
                                && change.getMemberChangesList().stream().anyMatch(
                                        member -> member.getId().equals(userId) && !member.getRemoved()))
                        .findFirst().orElse(null);
                if (seen == null) {
                    notifications.pause(Duration.ofSeconds(10));
                }
            }
            check(seen != null, "no directory pass showed the group with its new member");
            return "subscribed on try " + subscribed + "; notification " + notification.getChangeType()
                    + " for the group arrived (live: after " + waited + " s); directory pass " + passes
                    + " since the cursor shows the group with the new member";
        } finally {
            dropDirectoryFixtures();
        }
    }

    /** Subscribes to the run's group, which Graph may not know yet; answers the try that worked. */
    private int subscribeToTheNewGroup() throws Exception {
        CreateSubscriptionRequest request = CreateSubscriptionRequest.newBuilder()
                .setResource(SubscriptionResource.newBuilder().setGroup(entraGroupId))
                .addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED)
                .addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_DELETED)
                .setNotificationUrl(backend.notificationUrl())
                .setLifecycleNotificationUrl(backend.lifecycleUrl())
                .setExpirationDateTime(timestamp(backend.startedAt().plus(Duration.ofDays(1))))
                .build();
        StatusRuntimeException last = null;
        for (int attempt = 1; attempt <= DIRECTORY_TRIES; attempt++) {
            try {
                directorySubscriptionId = blocking.createSubscription(request).getSubscription().getId();
                check(!directorySubscriptionId.isBlank(), "Graph answered no subscription id");
                return attempt;
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != Status.Code.NOT_FOUND
                        && e.getStatus().getCode() != Status.Code.INVALID_ARGUMENT) {
                    throw e;
                }
                last = e;
                notifications.pause(Duration.ofSeconds(5));
            }
        }
        throw last;
    }

    /** Adds the user to the run's group, which Graph may not know yet. */
    private void addMember(String userId) throws Exception {
        StatusRuntimeException last = null;
        for (int attempt = 1; attempt <= DIRECTORY_TRIES; attempt++) {
            try {
                microsoft.addGroupMember(ai.pipestream.microsoft.v1.AddGroupMemberRequest.newBuilder()
                        .setGroupId(entraGroupId).setMemberId(userId).build());
                return;
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != Status.Code.NOT_FOUND) {
                    throw e;
                }
                last = e;
                notifications.pause(Duration.ofSeconds(5));
            }
        }
        throw last;
    }

    /** Deletes the subscription and the group of the directory step. */
    private void dropDirectoryFixtures() {
        if (directorySubscriptionId != null) {
            try {
                blocking.deleteSubscription(DeleteSubscriptionRequest.newBuilder()
                        .setSubscriptionId(directorySubscriptionId).build());
                directorySubscriptionId = null;
            } catch (RuntimeException e) {
                System.out.println("[ChangeScenarios] directory subscription clean-up failed: "
                        + e.getMessage());
            }
        }
        for (int attempt = 1; entraGroupId != null && attempt <= DIRECTORY_TRIES; attempt++) {
            try {
                microsoft.deleteGroup(ai.pipestream.microsoft.v1.DeleteGroupRequest.newBuilder()
                        .setGroupId(entraGroupId).build());
                entraGroupId = null;
            } catch (StatusRuntimeException e) {
                if (e.getStatus().getCode() != Status.Code.NOT_FOUND || attempt == DIRECTORY_TRIES) {
                    System.out.println("[ChangeScenarios] Entra group clean-up failed: " + e.getMessage());
                    return;
                }
                try {
                    notifications.pause(Duration.ofSeconds(5));
                } catch (InterruptedException interrupted) {
                    Thread.currentThread().interrupt();
                    return;
                } catch (Exception failed) {
                    return;
                }
            }
        }
    }

    private String subscribeToAUser() throws Exception {
        String id = graph.get("/users?$top=1&$select=id").path("value").path(0).path("id").asText("");
        check(!id.isBlank(), "the tenant lists no user");
        return subscribeTo(SubscriptionResource.newBuilder().setUser(id).build(), true);
    }

    private String subscribeToAGroup() throws Exception {
        return subscribeTo(SubscriptionResource.newBuilder().setGroup(firstGroup()).build(), true);
    }

    private String subscribeToGroupMembers() throws Exception {
        return subscribeTo(SubscriptionResource.newBuilder().setGroupMembers(firstGroup()).build(),
                false);
    }

    private String firstGroup() throws Exception {
        String id = graph.get("/groups?$top=1&$select=id").path("value").path(0).path("id").asText("");
        if (id.isBlank()) {
            throw new Skip("the tenant lists no group");
        }
        return id;
    }

    /** Creates and deletes a subscription, or reports that Graph refuses it to this app. */
    private String subscribeTo(SubscriptionResource resource, boolean deletes) {
        CreateSubscriptionRequest.Builder request = CreateSubscriptionRequest.newBuilder()
                .setResource(resource)
                .addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_UPDATED)
                .setNotificationUrl(backend.notificationUrl())
                .setLifecycleNotificationUrl(backend.lifecycleUrl())
                .setExpirationDateTime(timestamp(backend.startedAt().plus(Duration.ofDays(1))));
        if (deletes) {
            request.addChangeTypes(SubscriptionChangeType.SUBSCRIPTION_CHANGE_TYPE_DELETED);
        }
        Subscription created;
        try {
            created = blocking.createSubscription(request.build()).getSubscription();
        } catch (StatusRuntimeException e) {
            if (e.getStatus().getCode() == Status.Code.PERMISSION_DENIED) {
                return "REFUSED: Graph answered PERMISSION_DENIED to this app";
            }
            throw e;
        }
        check(!created.getId().isBlank(), "Graph answered no subscription id");
        blocking.deleteSubscription(DeleteSubscriptionRequest.newBuilder()
                .setSubscriptionId(created.getId()).build());
        return "created (" + created.getChangeTypesCount() + " change type(s)) and deleted";
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
        dropDirectoryFixtures();
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
        return awaitNotification(wait, subscriptionId);
    }

    private ChangeNotification awaitNotification(Duration wait, String ofSubscription)
            throws InterruptedException {
        long deadline = System.nanoTime() + wait.toNanos();
        while (System.nanoTime() < deadline) {
            ChangeNotification next = received.poll(500, TimeUnit.MILLISECONDS);
            if (next != null && next.getSubscriptionId().equals(ofSubscription)) {
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
