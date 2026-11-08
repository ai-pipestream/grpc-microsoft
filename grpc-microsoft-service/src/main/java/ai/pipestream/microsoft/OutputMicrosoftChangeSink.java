package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
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
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * Writes Graph crawl artifacts through the output SPI (protobuf, JSON,
 * OKF, microsoft-connector) into the selected store.
 *
 * <p>Per-record formats write on {@link #emit}. End-of-run formats (OKF) get a
 * mirror: the latest record per object key. A full-sweep run replaces the
 * mirror with exactly what it emitted; an incremental run merges its records
 * into the mirror, so unchanged entities stay in the bundle. An incremental run
 * in a process that has not completed a full sweep yet has no mirror to merge
 * into, so it writes its records under {@code {prefix}/increments/{run-id}}
 * instead of overwriting the full bundle with a partial one.</p>
 */
public final class OutputMicrosoftChangeSink implements MicrosoftChangeSink, AutoCloseable {

    private static final System.Logger LOG = System.getLogger(OutputMicrosoftChangeSink.class.getName());

    private final OutputStore store;
    private final List<OutputFormat> formats;
    private final String prefix;
    private final Map<String, Message> run = new LinkedHashMap<>();
    private final Map<String, Message> mirror = new LinkedHashMap<>();
    private boolean baseline;

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
        String relative = ObjectKeys.of(change);
        synchronized (this) {
            run.put(relative, change);
        }
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

    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        // Snapshots are crawl markers; artifacts come from emit/complete.
    }

    @Override
    public void completeRun(String runId) {
        completeRun(runId, true);
    }

    @Override
    public synchronized void completeRun(String runId, boolean fullSweep) {
        List<Message> records;
        String target = prefix;
        if (fullSweep) {
            mirror.clear();
            mirror.putAll(run);
            baseline = true;
            records = new ArrayList<>(mirror.values());
        } else if (baseline) {
            mirror.putAll(run);
            records = new ArrayList<>(mirror.values());
        } else {
            records = new ArrayList<>(run.values());
            target = ObjectKeys.under(prefix, ObjectKeys.join("increments", ObjectKeys.segment(runId)));
        }
        run.clear();
        try {
            for (OutputFormat format : formats) {
                format.complete(store, records, target);
            }
            LOG.log(System.Logger.Level.INFO,
                    "microsoft-proxy output wrote {0} records store={1} prefix={2} fullSweep={3}",
                    records.size(), store.id(), target, fullSweep);
        } catch (IOException e) {
            throw new UncheckedIOException("output complete failed for run " + runId, e);
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
