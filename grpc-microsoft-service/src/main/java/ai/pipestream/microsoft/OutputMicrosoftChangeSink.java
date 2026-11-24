package ai.pipestream.microsoft;

import ai.pipestream.microsoft.output.MicrosoftObjectKeys;
import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.output.ObjectKeys;
import ai.pipestream.output.OutputEnv;
import ai.pipestream.output.OutputFormat;
import ai.pipestream.output.OutputFormats;
import ai.pipestream.output.OutputStore;
import ai.pipestream.output.OutputStores;
import com.google.protobuf.Message;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Writes Graph crawl artifacts through the output SPI (protobuf, JSON,
 * OKF, microsoft-connector) into the selected store.
 *
 * <p>Per-record formats write on {@link #emit}. End-of-run formats (OKF) get a
 * mirror: the latest record per object key. A run replaces the mirror's
 * records of the kinds of entity it enumerated in full with exactly what it
 * emitted, and keeps the records of every other kind; an incremental run
 * merges its records into the mirror, so unchanged entities stay in the
 * bundle. An incremental run in a process that has not enumerated anything
 * in full yet has no mirror to merge into, so it writes its records under
 * {@code {prefix}/increments/{run-id}} instead of overwriting the full bundle
 * with a partial one (see {@link RunMirror}).</p>
 */
public final class OutputMicrosoftChangeSink implements MicrosoftChangeSink, AutoCloseable {

    private static final System.Logger LOG = System.getLogger(OutputMicrosoftChangeSink.class.getName());

    private final OutputStore store;
    private final List<OutputFormat> formats;
    private final String prefix;
    private final RunMirror<Message> mirror = new RunMirror<>();
    private final RunMirror<Message>.Pass own = mirror.pass("");

    /**
     * Creates a sink.
     *
     * @param store opened store
     * @param formats selected formats
     * @param prefix store prefix
     */
    public OutputMicrosoftChangeSink(OutputStore store, List<OutputFormat> formats, String prefix) {
        this.store = Objects.requireNonNull(store, "store");
        this.formats = List.copyOf(formats);
        this.prefix = prefix == null ? "" : prefix;
    }

    /**
     * Whether a store destination is configured.
     *
     * @return true when a directory or S3 bucket is set
     */
    public static boolean enabled() {
        return OutputEnv.destinationSet(OutputEnv.process());
    }

    /**
     * Builds a sink from the process environment.
     *
     * @return the sink
     */
    public static OutputMicrosoftChangeSink fromEnvironment() {
        return from(OutputEnv.process(), OutputStores.load(), OutputFormats.load());
    }

    /**
     * Builds a sink from {@code env} and catalogs.
     *
     * @param env environment
     * @param stores store catalog
     * @param formats format catalog
     * @return the sink
     */
    public static OutputMicrosoftChangeSink from(Map<String, String> env, OutputStores stores,
            OutputFormats formats) {
        try {
            OutputStore store = stores.select(env);
            List<OutputFormat> selected = formats.select(env);
            if (selected.isEmpty()) {
                throw new IllegalStateException("no output formats loaded for "
                        + OutputEnv.formats(env) + "; loaded=" + formats.ids());
            }
            LOG.log(System.Logger.Level.INFO,
                    "microsoft-proxy output store={0} loadedStores={1} formats={2}",
                    store.id(), stores.ids(), formats.ids());
            return new OutputMicrosoftChangeSink(store, selected, OutputEnv.prefix(env));
        } catch (IOException e) {
            throw new UncheckedIOException("output store failed to open", e);
        }
    }

    @Override
    public void emit(MicrosoftChange change) {
        emit(own, change);
    }

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        // Snapshots are crawl markers; artifacts come from emit/complete.
    }

    @Override
    public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated) {
        complete(own, runId, enumerated);
    }

    /**
     * This sink for the Sync passes of one connection. The view collects
     * that Sync's records on its own, and a pass of it replaces only that
     * connection's records in the bundle.
     *
     * @param connectionId the catalog connection id
     * @return the view; one per Sync
     */
    @Override
    public MicrosoftChangeSink forConnection(String connectionId) {
        RunMirror<Message>.Pass pass = mirror.pass(connectionId);
        return new MicrosoftChangeSink() {
            @Override
            public void emit(MicrosoftChange change) {
                OutputMicrosoftChangeSink.this.emit(pass, change);
            }

            @Override
            public void snapshot(MicrosoftSnapshot snapshot) {
                OutputMicrosoftChangeSink.this.snapshot(snapshot);
            }

            @Override
            public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated) {
                complete(pass, runId, enumerated);
            }
        };
    }

    private void emit(RunMirror<Message>.Pass pass, MicrosoftChange change) {
        String relative = MicrosoftObjectKeys.of(change);
        pass.put(relative, change.getEntity().getEntityCase(), change);
        String key = ObjectKeys.under(prefix, relative);
        try {
            for (OutputFormat format : formats) {
                if (format.supports(change)) {
                    format.emit(store, change, key);
                }
            }
        } catch (IOException e) {
            throw new UncheckedIOException("output emit failed for " + key, e);
        }
    }

    private void complete(RunMirror<Message>.Pass pass, String runId,
            Set<MicrosoftEntity.EntityCase> enumerated) {
        // One pass at a time writes the end-of-run artifacts.
        synchronized (mirror) {
            RunMirror.Completed<Message> completed = pass.complete(enumerated);
            String target = completed.increment()
                    ? ObjectKeys.under(prefix, ObjectKeys.join("increments",
                            pass.connectionId().isEmpty()
                                    ? "" : ObjectKeys.segment(pass.connectionId()),
                            ObjectKeys.segment(runId)))
                    : prefix;
            try {
                for (OutputFormat format : formats) {
                    format.complete(store, completed.values(), target);
                }
                LOG.log(System.Logger.Level.INFO, "microsoft-proxy output wrote {0} records "
                        + "store={1} prefix={2} connection={3} enumerated={4}",
                        completed.values().size(), store.id(), target, pass.connectionId(),
                        enumerated);
            } catch (IOException e) {
                throw new UncheckedIOException("output complete failed for run " + runId, e);
            }
        }
    }

    /**
     * Closes the store.
     */
    @Override
    public void close() {
        store.close();
    }
}
