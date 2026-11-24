package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.ChangeNotification;

import java.util.Set;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.LinkedBlockingQueue;

/** Fans accepted change notifications out to every open listener. */
final class NotificationBus {

    /** One listener's queue of notifications. */
    static final class Listener implements AutoCloseable {
        private final NotificationBus bus;
        private final BlockingQueue<ChangeNotification> queue = new LinkedBlockingQueue<>();

        private Listener(NotificationBus bus) {
            this.bus = bus;
        }

        /** The queue notifications arrive on. */
        BlockingQueue<ChangeNotification> queue() {
            return queue;
        }

        @Override
        public void close() {
            bus.listeners.remove(this);
        }
    }

    private final Set<Listener> listeners = ConcurrentHashMap.newKeySet();

    /** Opens a listener; close it when done. */
    Listener listen() {
        Listener listener = new Listener(this);
        listeners.add(listener);
        return listener;
    }

    /** Delivers one notification to every listener. */
    void publish(ChangeNotification notification) {
        for (Listener listener : listeners) {
            listener.queue.add(notification);
        }
    }
}
