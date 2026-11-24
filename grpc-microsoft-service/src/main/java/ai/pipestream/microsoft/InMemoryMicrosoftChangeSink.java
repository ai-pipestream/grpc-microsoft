package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.List;
import java.util.Set;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * A {@link MicrosoftChangeSink} that collects everything in memory, for tests
 * and for callers that want to batch-drain a crawl. Thread-safe by construction.
 */
public final class InMemoryMicrosoftChangeSink implements MicrosoftChangeSink {

    private final List<MicrosoftChange> changes = new CopyOnWriteArrayList<>();
    private final List<MicrosoftSnapshot> snapshots = new CopyOnWriteArrayList<>();
    private final List<String> completedRuns = new CopyOnWriteArrayList<>();
    private final List<Set<MicrosoftEntity.EntityCase>> enumeratedKinds =
            new CopyOnWriteArrayList<>();

    /** Creates an empty collector. */
    public InMemoryMicrosoftChangeSink() {
    }

    @Override
    public void emit(MicrosoftChange change) {
        changes.add(change);
    }

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        snapshots.add(snapshot);
    }

    /**
     * The changes collected so far, in emission order.
     *
     * @return an immutable copy
     */
    public List<MicrosoftChange> changes() {
        return List.copyOf(changes);
    }

    /**
     * The snapshots collected so far, in emission order.
     *
     * @return an immutable copy
     */
    public List<MicrosoftSnapshot> snapshots() {
        return List.copyOf(snapshots);
    }

    @Override
    public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated) {
        completedRuns.add(runId);
        enumeratedKinds.add(Set.copyOf(enumerated));
    }

    /**
     * For each entry of {@link #completedRuns()}, whether that run enumerated
     * any kind of entity in full.
     *
     * @return an immutable copy, in call order
     */
    public List<Boolean> fullSweeps() {
        return enumeratedKinds.stream().map(kinds -> !kinds.isEmpty()).toList();
    }

    /**
     * For each entry of {@link #completedRuns()}, the kinds of entity that
     * run enumerated in full.
     *
     * @return an immutable copy, in call order
     */
    public List<Set<MicrosoftEntity.EntityCase>> enumeratedKinds() {
        return List.copyOf(enumeratedKinds);
    }

    /**
     * Run ids passed to {@link #completeRun(String, Set)}, in call order.
     *
     * @return an immutable copy
     */
    public List<String> completedRuns() {
        return List.copyOf(completedRuns);
    }
}
