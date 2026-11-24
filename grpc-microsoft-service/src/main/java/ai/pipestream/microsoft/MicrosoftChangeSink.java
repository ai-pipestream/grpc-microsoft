package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

/**
 * Where the Microsoft crawler's output goes. Implementations must be
 * thread-safe: the crawler may emit from virtual threads. Optional
 * {@link OkfMicrosoftChangeSink} writes OKF + WARC when {@code OKF_DIR}
 * or {@code OKF_SPO_DRIVE_ID} is set.
 *
 * <p>A pass ends with {@link #completeRun(String, boolean)}. Only a complete
 * pass ({@code fullSweep}: every sweep was a full enumeration and the pass
 * covered everything the connection covers) lets a sink treat entities it
 * holds but the pass did not emit as gone. After an incremental (delta)
 * pass unseen entities are unchanged, and after a pass narrowed to some
 * drives, some sites or a folder they were never looked at; both must be
 * kept.</p>
 */
public interface MicrosoftChangeSink {

    /**
     * One upsert or delete against the Microsoft Graph mirror.
     *
     * @param change the change to emit
     */
    void emit(MicrosoftChange change);

    /**
     * The marker for one completed drive or list sweep; its {@code source} says
     * whether the sweep was a full enumeration.
     *
     * @param snapshot the snapshot to emit
     */
    void snapshot(MicrosoftSnapshot snapshot);

    /**
     * A full-sweep run finished. Sync-table sinks reconcile rows not seen in
     * {@code runId}; others ignore.
     *
     * @param runId identifier of the completed run
     */
    default void completeRun(String runId) {
    }

    /**
     * A run finished. The default forwards complete passes to
     * {@link #completeRun(String)} and ignores incremental and narrowed
     * passes, so a sink that reconciles by run id never deletes entities
     * that are unchanged or were outside the pass.
     *
     * @param runId identifier of the completed run
     * @param fullSweep whether the pass was complete: every sweep a full
     *        enumeration, over everything the connection covers
     *        ({@link MicrosoftCrawler.Result#complete()})
     */
    default void completeRun(String runId, boolean fullSweep) {
        if (fullSweep) {
            completeRun(runId);
        }
    }

    /**
     * This sink bound to a catalog connection, for a Sync that names one.
     *
     * @param connectionId the catalog connection id
     * @return a sink that stamps rows with {@code connectionId}; this sink when
     *         it has no per-connection state
     */
    default MicrosoftChangeSink forConnection(String connectionId) {
        return this;
    }
}
