package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftEntity;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * The latest value per key across runs, for a sink that writes a whole
 * bundle at the end of each run.
 *
 * <p>A run that enumerated some kinds of entity in full replaces the
 * mirror's entries of those kinds with exactly what it emitted, and keeps
 * the entries of every other kind: the run did not look at them. An
 * incremental or narrowed run enumerated nothing in full and merges its
 * values into the mirror, so unchanged entities stay in the bundle. Before
 * any run of this process has enumerated anything in full there is no
 * mirror to merge into, and such a run stands alone as an increment.</p>
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

    private record Held<V>(MicrosoftEntity.EntityCase kind, V value) {
    }

    private final Map<String, Held<V>> run = new LinkedHashMap<>();
    private final Map<String, Held<V>> mirror = new LinkedHashMap<>();
    private boolean baseline;

    /**
     * Records a value the current run emitted.
     *
     * @param key the entity's key in the bundle
     * @param kind the kind of entity
     * @param value the value to write for it
     */
    synchronized void put(String key, MicrosoftEntity.EntityCase kind, V value) {
        run.put(key, new Held<>(kind, value));
    }

    /**
     * Ends the current run.
     *
     * @param enumerated the kinds of entity the run enumerated in full
     * @return what to write
     */
    synchronized Completed<V> complete(Set<MicrosoftEntity.EntityCase> enumerated) {
        Map<String, Held<V>> source = mirror;
        boolean increment = false;
        if (!enumerated.isEmpty()) {
            mirror.values().removeIf(held -> enumerated.contains(held.kind()));
            mirror.putAll(run);
            baseline = true;
        } else if (baseline) {
            mirror.putAll(run);
        } else {
            source = run;
            increment = true;
        }
        List<V> values = new ArrayList<>(source.size());
        for (Held<V> held : source.values()) {
            values.add(held.value());
        }
        run.clear();
        return new Completed<>(List.copyOf(values), increment);
    }
}
