package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;

/**
 * A {@link MicrosoftChangeSink} that collects everything in memory, for tests
 * and for callers that want to batch-drain a crawl. Thread-safe by construction.
 */
public final class InMemoryMicrosoftChangeSink implements MicrosoftChangeSink {

    private final List<MicrosoftChange> changes = new CopyOnWriteArrayList<>();
    private final List<MicrosoftSnapshot> snapshots = new CopyOnWriteArrayList<>();
    private final List<String> completedRuns = new CopyOnWriteArrayList<>();

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
    public void completeRun(String runId) {
        completedRuns.add(runId);
    }

    /**
     * Run ids passed to {@link #completeRun(String)}, in call order.
     *
     * @return an immutable copy
     */
    public List<String> completedRuns() {
        return List.copyOf(completedRuns);
    }
}
