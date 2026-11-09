package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

/**
 * Where the Microsoft crawler's output goes. Implementations must be
 * thread-safe: the crawler may emit from virtual threads. Optional
 * {@link OkfMicrosoftChangeSink} writes OKF + WARC when {@code OKF_DIR}
 * or {@code OKF_SPO_DRIVE_ID} is set.
 *
 * <p>A pass ends with {@link #completeRun(String, boolean)}. Only a pass whose
 * every sweep was a full enumeration ({@code fullSweep}) lets a sink treat
 * entities it holds but the pass did not emit as gone; after an incremental
 * (delta) pass unseen entities are unchanged and must be kept.</p>
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
     * A run finished. The default forwards full sweeps to
     * {@link #completeRun(String)} and ignores incremental passes, so a sink that
     * reconciles by run id never deletes unchanged entities.
     *
     * @param runId identifier of the completed run
     * @param fullSweep whether every sweep of the run was a full enumeration
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
