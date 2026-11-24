package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftEntity;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * The latest value per key across runs, for a sink that writes a whole
 * bundle at the end of each run. The process has one such sink for every
 * connection, so each {@link Pass} collects the values of one crawl pass of
 * one connection and folds them into the mirror when it completes.
 *
 * <p>A run that enumerated some kinds of entity in full replaces its own
 * connection's entries of those kinds with exactly what it emitted. It
 * keeps the entries of every other kind, which it did not look at, and the
 * entries of every other connection, which it could not see. An incremental
 * or narrowed run enumerated nothing in full and merges its values into the
 * mirror, so unchanged entities stay in the bundle. Before any run of its
 * connection has enumerated anything in full there is nothing of that
 * connection to merge into, and such a run stands alone as an increment.</p>
 *
 * <p>Two passes running at once keep their own values until each completes,
 * so one Sync never writes, or discards, what another has emitted so far.</p>
 *
 * @param <V> what the sink writes per entity
 */
final class RunMirror<V> {

    /**
     * What a run leaves the sink to write.
     *
     * @param values the values of the bundle, or of the increment
     * @param increment true when {@code values} is only this run's own
     *        output, to be written beside the bundle and not over it
     * @param <V> what the sink writes per entity
     */
    record Completed<V>(List<V> values, boolean increment) {
    }

    private record Held<V>(String owner, MicrosoftEntity.EntityCase kind, V value) {
    }

    private final Map<String, Held<V>> mirror = new LinkedHashMap<>();
    /** Connections with a run that enumerated something in full. */
    private final Set<String> baselines = new HashSet<>();
    private final Pass own = new Pass("");

    /**
     * Starts collecting for {@code connectionId}. One pass object serves the
     * crawl passes of one Sync, never two at once.
     *
     * @param connectionId the connection whose entries the pass owns; empty
     *        for the process-default connection
     * @return the collector
     */
    Pass pass(String connectionId) {
        return new Pass(connectionId == null ? "" : connectionId);
    }

    /**
     * Records a value the current run of the process-default connection
     * emitted.
     *
     * @param key the entity's key in the bundle
     * @param kind the kind of entity
     * @param value the value to write for it
     */
    void put(String key, MicrosoftEntity.EntityCase kind, V value) {
        own.put(key, kind, value);
    }

    /**
     * Ends the current run of the process-default connection.
     *
     * @param enumerated the kinds of entity the run enumerated in full
     * @return what to write
     */
    Completed<V> complete(Set<MicrosoftEntity.EntityCase> enumerated) {
        return own.complete(enumerated);
    }

    /** The values of the crawl pass in progress for one connection. */
    final class Pass {

        private final String owner;
        private final Map<String, Held<V>> run = new LinkedHashMap<>();

        private Pass(String owner) {
            this.owner = owner;
        }

        /**
         * The connection this pass collects for.
         *
         * @return the connection id; empty for the process-default connection
         */
        String connectionId() {
            return owner;
        }

        /**
         * Records a value the pass emitted; the latest value for a key wins.
         *
         * @param key the entity's key in the bundle
         * @param kind the kind of entity
         * @param value the value to write for it
         */
        void put(String key, MicrosoftEntity.EntityCase kind, V value) {
            synchronized (run) {
                run.put(key, new Held<>(owner, kind, value));
            }
        }

        /**
         * Ends the pass in progress and folds it into the mirror.
         *
         * @param enumerated the kinds of entity the pass enumerated in full
         * @return what to write
         */
        Completed<V> complete(Set<MicrosoftEntity.EntityCase> enumerated) {
            Map<String, Held<V>> emitted;
            synchronized (run) {
                emitted = new LinkedHashMap<>(run);
                run.clear();
            }
            synchronized (RunMirror.this) {
                if (!enumerated.isEmpty()) {
                    mirror.values().removeIf(held -> held.owner().equals(owner)
                            && enumerated.contains(held.kind()));
                    baselines.add(owner);
                } else if (!baselines.contains(owner)) {
                    return new Completed<>(values(emitted), true);
                }
                mirror.putAll(emitted);
                return new Completed<>(values(mirror), false);
            }
        }
    }

    private static <V> List<V> values(Map<String, Held<V>> held) {
        List<V> values = new ArrayList<>(held.size());
        for (Held<V> one : held.values()) {
            values.add(one.value());
        }
        return List.copyOf(values);
    }
}
