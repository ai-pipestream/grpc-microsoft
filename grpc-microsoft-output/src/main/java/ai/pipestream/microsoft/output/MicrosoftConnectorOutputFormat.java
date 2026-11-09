package ai.pipestream.microsoft.output;

import ai.pipestream.microsoft.v1.MicrosoftChange;
import ai.pipestream.output.OutputFormat;
import ai.pipestream.output.OutputObject;
import ai.pipestream.output.OutputStore;
import com.google.protobuf.Message;

import java.io.IOException;

/**
 * Binary {@link MicrosoftChange} at {@code {key}.connector.pb} — the same
 * protobuf the Graph Connector Agent adapter consumes from
 * {@code MicrosoftService.Sync}. Confluence records are skipped.
 */
public final class MicrosoftConnectorOutputFormat implements OutputFormat {

    /** Format id. */
    public static final String ID = "microsoft-connector";
    /** Media type. */
    public static final String MEDIA_TYPE = "application/x-protobuf";

    /**
     * Creates the format (ServiceLoader).
     */
    public MicrosoftConnectorOutputFormat() {
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
    public void emit(OutputStore store, Message record, String key) throws IOException {
        if (record instanceof MicrosoftChange change) {
            store.put(OutputObject.of(key + ".connector.pb", change.toByteArray(), MEDIA_TYPE));
        }
    }
}
