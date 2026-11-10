package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.okf.CatalogEntry;
import ai.pipestream.okf.OkfCatalogSource;
import com.google.protobuf.Message;

import java.util.Optional;

/**
 * Registers {@link MicrosoftCatalog} with the {@code okf} output format.
 */
public final class MicrosoftOkfSource implements OkfCatalogSource {

    /** Source id. */
    public static final String ID = "microsoft";

    /**
     * Creates the source (ServiceLoader).
     */
    public MicrosoftOkfSource() {
    }

    @Override
    public String id() {
        return ID;
    }

    @Override
    public boolean supports(Message record) {
        return record instanceof MicrosoftChange;
    }

    @Override
    public Optional<CatalogEntry> entry(Message record) {
        return record instanceof MicrosoftChange change
                ? MicrosoftCatalog.from(change)
                : Optional.empty();
    }

    @Override
    public String title() {
        return "Microsoft Graph crawl";
    }

    @Override
    public String description() {
        return "OKF v0.2 capture of a MicrosoftService.Sync run.";
    }

    @Override
    public String actor() {
        return MicrosoftCatalog.ACTOR;
    }

    @Override
    public String software() {
        return "grpc-microsoft/okf-producer";
    }
}
