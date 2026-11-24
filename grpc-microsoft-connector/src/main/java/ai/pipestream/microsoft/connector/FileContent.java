package ai.pipestream.microsoft.connector;

import ai.pipestream.microsoft.v1.DriveItem;
import microsoft.graph.connectors.contracts.grpc.Content;

import java.nio.ByteBuffer;
import java.nio.charset.CharacterCodingException;
import java.nio.charset.Charset;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.IllegalCharsetNameException;
import java.nio.charset.StandardCharsets;
import java.nio.charset.UnsupportedCharsetException;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Turns a drive item's inlined bytes into GCA {@link Content}, or says why
 * it could not.
 *
 * <p>GCA content is text or HTML, so only text media types are sent. The
 * bytes are decoded strictly with the charset the file names: a byte order
 * mark first, then the media type's {@code charset} parameter, then the
 * declaration inside an HTML or XML document, and UTF-8 when nothing names
 * one. Bytes that are not valid in that charset are left out instead of
 * being sent with replacement characters.
 */
final class FileContent {

    /** Media types outside {@code text/} whose bytes are text. */
    private static final Set<String> TEXT_APPLICATION_TYPES = Set.of(
            "application/json",
            "application/xml",
            "application/javascript",
            "application/x-javascript",
            "application/yaml",
            "application/x-yaml",
            "application/ndjson",
            "application/x-ndjson");

    /** {@code text/} types that are markup of a binary-era format, not readable text. */
    private static final Set<String> NOT_PLAIN_TEXT_TYPES = Set.of("text/rtf", "text/richtext");

    private static final Set<String> HTML_TYPES = Set.of("text/html", "application/xhtml+xml");

    /** How far into a document its own charset declaration is looked for. */
    private static final int DECLARATION_SCAN_BYTES = 1024;

    private static final Pattern HTML_CHARSET = Pattern.compile(
            "<meta[^>]*charset\\s*=\\s*[\"']?\\s*([A-Za-z0-9._:-]+)", Pattern.CASE_INSENSITIVE);
    private static final Pattern XML_ENCODING = Pattern.compile(
            "^\\s*<\\?xml[^>]*encoding\\s*=\\s*[\"']([A-Za-z0-9._:-]+)[\"']");

    private FileContent() {
    }

    /**
     * The outcome for one drive item: content, an omission, or neither when
     * there was nothing to send.
     *
     * @param content the content to set on the GCA item
     * @param omitted why the file's content is not on the item
     */
    record Outcome(Optional<Content> content, Optional<OmittedContent> omitted) {

        private static final Outcome NOTHING = new Outcome(Optional.empty(), Optional.empty());

        private static Outcome omitted(OmittedContent.Reason reason, String detail) {
            return new Outcome(Optional.empty(), Optional.of(new OmittedContent(reason, detail)));
        }
    }

    /**
     * Maps {@code item}'s bytes to GCA content.
     *
     * @param item the drive item
     * @param contentRequested whether the crawl asked MicrosoftService to
     *        inline file bytes; when it did not, a file without bytes is
     *        not an omission
     * @return the content, the reason there is none, or neither for a
     *         folder, an empty file and a crawl that did not ask for content
     */
    static Outcome of(DriveItem item, boolean contentRequested) {
        boolean inlined = item.hasContent() && !item.getContent().isEmpty();
        if (item.getFolder() || (!inlined && (!contentRequested || item.getSize() == 0))) {
            return Outcome.NOTHING;
        }
        MediaType mediaType = MediaType.parse(item.getMimeType());
        if (!mediaType.isText()) {
            return Outcome.omitted(OmittedContent.Reason.NOT_TEXT, mediaType.type().isEmpty()
                    ? "the file has no media type"
                    : "media type " + mediaType.type() + " is not text");
        }
        long rawSize = inlined ? item.getContent().size() : item.getSize();
        if (rawSize > CrawlItemMapper.CONTENT_CAP_BYTES) {
            return tooLarge("the file is " + rawSize + " bytes");
        }
        if (!inlined) {
            return Outcome.omitted(OmittedContent.Reason.NOT_INLINED,
                    "MicrosoftService did not inline the file's " + item.getSize() + " bytes");
        }
        byte[] bytes = item.getContent().toByteArray();
        Bom bom = Bom.of(bytes);
        Charset charset;
        try {
            charset = bom != null ? bom.charset() : charset(mediaType, bytes);
        } catch (IllegalCharsetNameException | UnsupportedCharsetException e) {
            return Outcome.omitted(OmittedContent.Reason.NOT_DECODABLE,
                    "charset " + e.getMessage() + " is not supported");
        }
        int offset = bom == null ? 0 : bom.length();
        String text;
        try {
            text = charset.newDecoder()
                    .onMalformedInput(CodingErrorAction.REPORT)
                    .onUnmappableCharacter(CodingErrorAction.REPORT)
                    .decode(ByteBuffer.wrap(bytes, offset, bytes.length - offset))
                    .toString();
        } catch (CharacterCodingException e) {
            return Outcome.omitted(OmittedContent.Reason.NOT_DECODABLE,
                    "the bytes are not valid " + charset.name());
        }
        // The limit is on what goes on the wire, and GCA strings travel as UTF-8.
        int utf8Size = charset.equals(StandardCharsets.UTF_8)
                ? bytes.length - offset
                : text.getBytes(StandardCharsets.UTF_8).length;
        if (utf8Size > CrawlItemMapper.CONTENT_CAP_BYTES) {
            return tooLarge("the text is " + utf8Size + " bytes as UTF-8");
        }
        return new Outcome(Optional.of(Content.newBuilder()
                .setContentType(mediaType.isHtml() ? Content.ContentType.Html
                        : Content.ContentType.Text)
                .setContentValue(text)
                .build()), Optional.empty());
    }

    private static Outcome tooLarge(String found) {
        return Outcome.omitted(OmittedContent.Reason.TOO_LARGE,
                found + ", over the " + CrawlItemMapper.CONTENT_CAP_BYTES + " byte limit");
    }

    /** The charset of a document without a byte order mark. */
    private static Charset charset(MediaType mediaType, byte[] bytes) {
        if (mediaType.charset().isPresent()) {
            return Charset.forName(mediaType.charset().get());
        }
        Pattern declaration = mediaType.isHtml() ? HTML_CHARSET
                : mediaType.isXml() ? XML_ENCODING : null;
        if (declaration != null) {
            String head = new String(bytes, 0, Math.min(bytes.length, DECLARATION_SCAN_BYTES),
                    StandardCharsets.ISO_8859_1);
            Matcher matcher = declaration.matcher(head);
            if (matcher.find()) {
                Charset declared = Charset.forName(matcher.group(1));
                // A declaration that could be read as single bytes is not in a
                // multi-byte-unit charset, whatever it says.
                String name = declared.name().toUpperCase(Locale.ROOT);
                return name.startsWith("UTF-16") || name.startsWith("UTF-32")
                        ? StandardCharsets.UTF_8 : declared;
            }
        }
        return StandardCharsets.UTF_8;
    }

    /** A media type without its parameters, and its {@code charset} parameter. */
    private record MediaType(String type, Optional<String> charset) {

        static MediaType parse(String raw) {
            String[] parts = raw.split(";");
            String type = parts[0].trim().toLowerCase(Locale.ROOT);
            Optional<String> charset = Optional.empty();
            for (int i = 1; i < parts.length; i++) {
                String parameter = parts[i].trim();
                int equals = parameter.indexOf('=');
                if (equals > 0 && parameter.substring(0, equals).trim()
                        .equalsIgnoreCase("charset")) {
                    String value = parameter.substring(equals + 1).trim();
                    if (value.length() >= 2 && value.startsWith("\"") && value.endsWith("\"")) {
                        value = value.substring(1, value.length() - 1);
                    }
                    if (!value.isEmpty()) {
                        charset = Optional.of(value);
                    }
                }
            }
            return new MediaType(type, charset);
        }

        boolean isHtml() {
            return HTML_TYPES.contains(type);
        }

        boolean isXml() {
            return type.equals("text/xml") || type.equals("application/xml")
                    || type.endsWith("+xml");
        }

        boolean isText() {
            if (NOT_PLAIN_TEXT_TYPES.contains(type)) {
                return false;
            }
            return type.startsWith("text/") || TEXT_APPLICATION_TYPES.contains(type)
                    || type.endsWith("+json") || type.endsWith("+xml");
        }
    }

    /** A byte order mark: the charset it names and how many bytes it takes. */
    private record Bom(Charset charset, int length) {

        static Bom of(byte[] b) {
            if (starts(b, 0xEF, 0xBB, 0xBF)) {
                return new Bom(StandardCharsets.UTF_8, 3);
            }
            // The UTF-32 marks begin with the UTF-16 ones, so they go first.
            if (starts(b, 0x00, 0x00, 0xFE, 0xFF)) {
                return new Bom(Charset.forName("UTF-32BE"), 4);
            }
            if (starts(b, 0xFF, 0xFE, 0x00, 0x00)) {
                return new Bom(Charset.forName("UTF-32LE"), 4);
            }
            if (starts(b, 0xFE, 0xFF)) {
                return new Bom(StandardCharsets.UTF_16BE, 2);
            }
            if (starts(b, 0xFF, 0xFE)) {
                return new Bom(StandardCharsets.UTF_16LE, 2);
            }
            return null;
        }

        private static boolean starts(byte[] bytes, int... mark) {
            if (bytes.length < mark.length) {
                return false;
            }
            for (int i = 0; i < mark.length; i++) {
                if ((bytes[i] & 0xFF) != mark[i]) {
                    return false;
                }
            }
            return true;
        }
    }
}
