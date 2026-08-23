package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.microsoft.v1.MicrosoftSnapshot;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.KnowledgeBundle;
import ai.pipestream.okf.OkfOutput;
import ai.pipestream.okf.microsoft.MicrosoftCatalog;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Accumulates Microsoft Graph changes into an OKF v0.2 directory tree, zip,
 * and sibling WARC 1.1 file. Optionally uploads that payload to a SharePoint
 * library ({@code OKF_SPO_DRIVE_ID}). Materializes on
 * {@link #completeRun(String)}.
 */
public final class OkfMicrosoftChangeSink implements MicrosoftChangeSink {

    private static final System.Logger LOG = System.getLogger(OkfMicrosoftChangeSink.class.getName());

    private final OkfOutput output;
    private final SharePointOkfPublisher publisher;
    private final Map<String, CatalogEntry> entries = new ConcurrentHashMap<>();

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
        MicrosoftCatalog.from(change).ifPresent(entry -> entries.put(entry.path(), entry));
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
    public void completeRun(String runId) {
        List<CatalogEntry> captured = new ArrayList<>(entries.values());
        Path scratch = null;
        try {
            OkfOutput dest = output;
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
                publisher.publish(bundle.okf(), dest);
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
