package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;

import java.util.Collections;
import java.util.EnumSet;
import java.util.Set;

/**
 * Where the Microsoft crawler's output goes. Implementations must be
 * thread-safe: the crawler may emit from virtual threads. Optional
 * {@link OkfMicrosoftChangeSink} writes OKF + WARC when {@code OKF_DIR}
 * or {@code OKF_SPO_DRIVE_ID} is set.
 *
 * <p>A pass ends with {@link #completeRun(String, Set, Set)}, which names
 * the kinds of entity the pass enumerated in full over everything the
 * connection covers, and the single drives and lists it enumerated in full
 * when it did not cover the whole connection. Only for those kinds, and
 * inside those drives and lists, may a sink treat an entity it holds but
 * the pass did not emit as gone. After an incremental (delta) pass unseen
 * entities are unchanged; after a pass narrowed to some drives, some sites
 * or a folder they were never looked at; and a pass that did not ask for
 * lists or for security never read those kinds at all. In each case the
 * entities must be kept.</p>
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
     * Every kind of entity a change can carry.
     */
    Set<MicrosoftEntity.EntityCase> ALL_KINDS = Collections.unmodifiableSet(
            EnumSet.complementOf(EnumSet.of(MicrosoftEntity.EntityCase.ENTITY_NOT_SET)));

    /**
     * A run that enumerated every kind of entity in full finished. Forwards
     * to {@link #completeRun(String, Set)} with {@link #ALL_KINDS}.
     *
     * @param runId identifier of the completed run
     */
    default void completeRun(String runId) {
        completeRun(runId, ALL_KINDS);
    }

    /**
     * A run finished. Forwards to {@link #completeRun(String, Set)} with
     * {@link #ALL_KINDS} for a complete pass and with no kind otherwise.
     *
     * @param runId identifier of the completed run
     * @param fullSweep whether the pass enumerated every kind of entity in
     *        full, over everything the connection covers
     */
    default void completeRun(String runId, boolean fullSweep) {
        completeRun(runId, fullSweep ? ALL_KINDS : Set.of());
    }

    /**
     * A run that enumerated no single drive or list on its own finished.
     * Forwards to {@link #completeRun(String, Set, Set)} with no scope.
     *
     * @param runId identifier of the completed run
     * @param enumerated the kinds of entity the pass enumerated in full over
     *        everything the connection covers
     */
    default void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated) {
        completeRun(runId, enumerated, Set.of());
    }

    /**
     * A run finished. Sync-table sinks reconcile, for each kind in
     * {@code enumerated}, the rows the run did not write, and for each scope
     * in {@code scopes} the rows of that drive or list the run did not
     * write; mirrors drop the same entities. Everything else is kept. The
     * default does nothing.
     *
     * @param runId identifier of the completed run
     * @param enumerated the kinds of entity the pass enumerated in full over
     *        everything the connection covers
     *        ({@link MicrosoftCrawler.Result#enumerated()}); empty after an
     *        incremental or a narrowed pass
     * @param scopes the drives and lists the pass enumerated in full on
     *        their own ({@link MicrosoftCrawler.Result#scopes()}); empty
     *        when {@code enumerated} already covers them
     */
    default void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated,
            Set<SweepScope> scopes) {
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
