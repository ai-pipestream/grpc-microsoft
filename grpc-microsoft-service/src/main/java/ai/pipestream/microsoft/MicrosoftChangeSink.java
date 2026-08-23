package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

/**
 * Where the Microsoft crawler's output goes. Implementations must be
 * thread-safe: the crawler may emit from virtual threads. Optional
 * {@link OkfMicrosoftChangeSink} writes OKF + WARC when {@code OKF_DIR}
 * or {@code OKF_SPO_DRIVE_ID} is set.
 */
public interface MicrosoftChangeSink {

    /**
     * One upsert or delete against the Microsoft Graph mirror.
     *
     * @param change the change to emit
     */
    void emit(MicrosoftChange change);

    /**
     * The full-sync marker for one completed drive crawl.
     *
     * @param snapshot the snapshot to emit
     */
    void snapshot(MicrosoftSnapshot snapshot);

    /**
     * A full crawl run finished. Sync-table sinks reconcile rows not seen
     * in {@code runId}; others ignore.
     *
     * @param runId identifier of the completed run
     */
    default void completeRun(String runId) {
    }
}
