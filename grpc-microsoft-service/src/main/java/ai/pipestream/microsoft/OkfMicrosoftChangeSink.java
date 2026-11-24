package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftEntity;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.KnowledgeBundle;
import ai.pipestream.okf.OkfOutput;
import ai.pipestream.microsoft.output.MicrosoftCatalog;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.List;
import java.util.Set;

/**
 * Accumulates Microsoft Graph changes into an OKF v0.2 directory tree, zip,
 * and sibling WARC 1.1 file. Optionally uploads that payload to a SharePoint
 * library ({@code OKF_SPO_DRIVE_ID}). Materializes on
 * {@link #completeRun(String, Set, Set)}.
 *
 * <p>The bundle is a mirror: a run replaces the entries of the kinds of
 * entity it enumerated in full with exactly the entries it emitted and keeps
 * the entries of every other kind, an incremental run merges its entries
 * (deletes become {@code status: deprecated}) into the mirror, so unchanged
 * entities are never dropped. An incremental run with nothing enumerated in
 * full earlier in this process writes only its own entries, beside the
 * configured bundle under an {@code increments} name, never over it (see
 * {@link RunMirror}).</p>
 */
public final class OkfMicrosoftChangeSink implements MicrosoftChangeSink {

    private static final System.Logger LOG = System.getLogger(OkfMicrosoftChangeSink.class.getName());

    private final OkfOutput output;
    private final SharePointOkfPublisher publisher;
    private final RunMirror<CatalogEntry> mirror = new RunMirror<>();
    private final RunMirror<CatalogEntry>.Pass own = mirror.pass("");

    /**
     * Creates a sink.
     *
     * @param output local destinations; may be empty when only SharePoint is set
     * @param publisher SharePoint uploader, or {@code null} for disk only
     */
    public OkfMicrosoftChangeSink(OkfOutput output, SharePointOkfPublisher publisher) {
        this.output = output;
        this.publisher = publisher;
    }

    /**
     * Whether disk OKF destinations or SharePoint upload are configured.
     *
     * @return true when the sink should be wired
     */
    public static boolean enabled() {
        return OkfOutput.fromEnvironment().enabled() || SharePointOkfPublisher.enabled();
    }

    /**
     * Builds a sink from the process environment.
     *
     * @param files Graph files API (used only when SharePoint upload is on)
     * @return the sink
     */
    public static OkfMicrosoftChangeSink fromEnvironment(GraphFiles files) {
        SharePointOkfPublisher spo = SharePointOkfPublisher.enabled()
                ? SharePointOkfPublisher.fromEnvironment(files) : null;
        return new OkfMicrosoftChangeSink(OkfOutput.fromEnvironment(), spo);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void emit(MicrosoftChange change) {
        emit(own, change);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void snapshot(MicrosoftSnapshot snapshot) {
        // Snapshots are crawl markers; concepts come from emit().
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated,
            Set<SweepScope> scopes) {
        complete(own, runId, enumerated, scopes);
    }

    /**
     * This sink for the Sync passes of one connection. The view collects
     * that Sync's entries on its own, and a pass of it replaces only that
     * connection's entries in the bundle.
     *
     * @param connectionId the catalog connection id
     * @return the view; one per Sync
     */
    @Override
    public MicrosoftChangeSink forConnection(String connectionId) {
        RunMirror<CatalogEntry>.Pass pass = mirror.pass(connectionId);
        return new MicrosoftChangeSink() {
            @Override
            public void emit(MicrosoftChange change) {
                OkfMicrosoftChangeSink.emit(pass, change);
            }

            @Override
            public void snapshot(MicrosoftSnapshot snapshot) {
                OkfMicrosoftChangeSink.this.snapshot(snapshot);
            }

            @Override
            public void completeRun(String runId, Set<MicrosoftEntity.EntityCase> enumerated,
                    Set<SweepScope> scopes) {
                complete(pass, runId, enumerated, scopes);
            }
        };
    }

    private static void emit(RunMirror<CatalogEntry>.Pass pass, MicrosoftChange change) {
        MicrosoftCatalog.from(change).ifPresent(entry ->
                pass.put(entry.path(), change.getEntity().getEntityCase(),
                        SweepScope.idOf(change.getEntity()), entry));
    }

    private void complete(RunMirror<CatalogEntry>.Pass pass, String runId,
            Set<MicrosoftEntity.EntityCase> enumerated, Set<SweepScope> scopes) {
        // One pass at a time writes the bundle.
        synchronized (mirror) {
            RunMirror.Completed<CatalogEntry> completed = pass.complete(enumerated, scopes);
            String connection = pass.connectionId().isEmpty() ? ""
                    : ai.pipestream.okf.OkfPaths.segment(pass.connectionId()) + "-";
            write(connection + ai.pipestream.okf.OkfPaths.segment(runId), runId,
                    completed.values(), completed.increment());
        }
    }

    /** Writes the bundle, or an increment named by {@code incrementName} beside it. */
    private void write(String incrementName, String runId, List<CatalogEntry> captured,
            boolean increment) {
        Path scratch = null;
        String suffix = "increments-" + incrementName;
        try {
            OkfOutput dest = increment ? beside(output, suffix) : output;
            if (!dest.enabled()) {
                scratch = Files.createTempDirectory("okf-microsoft-");
                dest = new OkfOutput(scratch, scratch.resolve("bundle.zip"),
                        scratch.resolve("bundle.warc.gz"));
            }
            KnowledgeBundle bundle = KnowledgeBundle.assemble(
                    "Microsoft Graph crawl",
                    "OKF v0.2 capture of a MicrosoftService.Sync run.",
                    MicrosoftCatalog.ACTOR,
                    Instant.now(),
                    "grpc-microsoft/okf-producer",
                    captured);
            bundle.write(dest);
            if (publisher != null) {
                (increment ? publisher.under("increments/" + incrementName)
                        : publisher).publish(bundle.okf(), dest);
            }
            LOG.log(System.Logger.Level.INFO,
                    "microsoft-proxy OKF wrote {0} concepts dir={1} zip={2} warc={3} spo={4}",
                    captured.size(), dest.directory(), dest.zip(), dest.warc(),
                    publisher != null);
        } catch (IOException e) {
            throw new UncheckedIOException("OKF write failed for run " + runId, e);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IllegalStateException("OKF SharePoint upload interrupted", e);
        } finally {
            if (scratch != null) {
                deleteQuietly(scratch);
            }
        }
    }

    /** The configured destinations renamed with {@code suffix}, beside the originals. */
    static OkfOutput beside(OkfOutput output, String suffix) {
        return new OkfOutput(
                output.directory() == null ? null
                        : output.directory().resolveSibling(name(output.directory()) + "-" + suffix),
                output.zip() == null ? null : output.zip().resolveSibling(
                        stem(name(output.zip()), ".zip") + "-" + suffix + ".zip"),
                output.warc() == null ? null : output.warc().resolveSibling(
                        stem(name(output.warc()), ".warc.gz") + "-" + suffix + ".warc.gz"));
    }

    private static String name(Path path) {
        return path.getFileName() == null ? "bundle" : path.getFileName().toString();
    }

    private static String stem(String name, String extension) {
        return name.endsWith(extension) ? name.substring(0, name.length() - extension.length()) : name;
    }

    private static void deleteQuietly(Path root) {
        try (var walk = Files.walk(root)) {
            walk.sorted((a, b) -> b.getNameCount() - a.getNameCount()).forEach(path -> {
                try {
                    Files.deleteIfExists(path);
                } catch (IOException ignored) {
                    // best-effort temp cleanup
                }
            });
        } catch (IOException ignored) {
            // best-effort temp cleanup
        }
    }
}
