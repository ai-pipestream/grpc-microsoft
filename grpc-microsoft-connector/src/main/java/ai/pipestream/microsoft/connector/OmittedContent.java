package ai.pipestream.microsoft.connector;

/**
 * Why a file's bytes did not become the content of its GCA item. The item
 * is still emitted with its properties; the crawl bit that carries it
 * reports {@code PartialSuccess} with {@link #detail()} in its message.
 *
 * @param reason the kind of omission
 * @param detail what was found, for the status message: the media type,
 *        the size and the limit, or the charset
 */
public record OmittedContent(Reason reason, String detail) {

    /** The kinds of omission. */
    public enum Reason {
        /** The media type is not one of the text types the connector sends. */
        NOT_TEXT,
        /** The file, or its text once decoded, is over the content limit. */
        TOO_LARGE,
        /** Content was asked for and MicrosoftService did not inline it. */
        NOT_INLINED,
        /** The bytes are not valid text in the file's charset. */
        NOT_DECODABLE
    }
}
