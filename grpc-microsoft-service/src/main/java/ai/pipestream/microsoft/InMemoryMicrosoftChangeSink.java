package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;

/** Collects crawl output in memory for tests. */
public final class InMemoryMicrosoftChangeSink implements MicrosoftChangeSink {

    private final List<MicrosoftChange> changes = new CopyOnWriteArrayList<>();
    private final List<MicrosoftSnapshot> snapshots = new CopyOnWriteArrayList<>();
    private final List<String> completedRuns = new CopyOnWriteArrayList<>();

    @Override
    public void emit(MicrosoftChange change) {
        changes.add(change);
    }

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        snapshots.add(snapshot);
    }

    public List<MicrosoftChange> changes() {
        return List.copyOf(changes);
    }

    public List<MicrosoftSnapshot> snapshots() {
        return List.copyOf(snapshots);
    }

    @Override
    public void completeRun(String runId) {
        completedRuns.add(runId);
    }

    public List<String> completedRuns() {
        return List.copyOf(completedRuns);
    }
}
