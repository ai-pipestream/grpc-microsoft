package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * Replaces tenant data in recorded HTTP exchanges with placeholders. One real
 * value always becomes the same placeholder, so ids still line up across a run.
 */
final class RecordingScrubber {

    /** Start of every placeholder GUID. */
    static final String GUID_PREFIX = "00000000-0000-4000-8000-";
    /** Start of every placeholder URL. */
    static final String URL_PREFIX = "https://redacted.example/url/";

    /** JSON objects whose {@code displayName} names a person, group or app. */
    private static final Set<String> IDENTITIES = Set.of("user", "application", "device", "group",
            "siteUser", "siteGroup", "sharePointGroup");
    /** Host suffixes that name the tenant. */
    private static final List<String> TENANT_HOSTS = List.of(".sharepoint.com",
            ".onmicrosoft.com", ".sharepoint.us", ".sharepoint.de", ".sharepoint.cn");

    private final String graphOrigin;
    private final Map<String, String> placeholders = new HashMap<>();
    private final Map<String, Integer> counters = new HashMap<>();

    /**
     * Creates a scrubber.
     *
     * @param graphBaseUrl the Graph base URL; URLs on its origin keep their path
     */
    RecordingScrubber(String graphBaseUrl) {
        this.graphOrigin = origin(graphBaseUrl);
    }

    /** Scrubs every string value of a JSON tree; keys stay. */
    synchronized JsonNode json(JsonNode node) {
        return json(node, "");
    }

    private JsonNode json(JsonNode node, String parentKey) {
        if (node.isObject()) {
            ObjectNode out = JsonNodeFactory.instance.objectNode();
            for (Map.Entry<String, JsonNode> entry : node.properties()) {
                JsonNode value = entry.getValue();
                if (entry.getKey().equals("displayName") && value.isTextual()
                        && IDENTITIES.contains(parentKey)) {
                    out.put(entry.getKey(), placeholder("name", value.asText(), "Identity %d"));
                } else {
                    out.set(entry.getKey(), json(value, entry.getKey()));
                }
            }
            return out;
        }
        if (node.isArray()) {
            ArrayNode out = JsonNodeFactory.instance.arrayNode();
            node.forEach(element -> out.add(json(element, parentKey)));
            return out;
        }
        return node.isTextual() ? JsonNodeFactory.instance.textNode(text(node.asText())) : node;
    }

    /** Scrubs free text: URLs, tokens, GUIDs, drive and item ids, hosts and emails. */
    synchronized String text(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }
        String guidless = guids(input);
        StringBuilder out = new StringBuilder(guidless.length());
        int i = 0;
        while (i < guidless.length()) {
            int url = urlEnd(guidless, i);
            if (url > i) {
                String value = guidless.substring(i, url);
                String origin = origin(value);
                if (origin.equalsIgnoreCase(graphOrigin)) {
                    out.append(graphOrigin);
                    i += origin.length();
                } else {
                    out.append(placeholder("url", value, URL_PREFIX + "%d"));
                    i = url;
                }
                continue;
            }
            char c = guidless.charAt(i);
            boolean boundary = i == 0 || !wordChar(guidless.charAt(i - 1));
            if (boundary && guidless.startsWith("eyJ", i)) {
                int end = run(guidless, i, RecordingScrubber::tokenChar);
                if (end - i >= 20) {
                    out.append("token-redacted");
                    i = end;
                    continue;
                }
            }
            if (boundary && guidless.startsWith("b!", i)) {
                int end = run(guidless, i + 2, RecordingScrubber::idChar);
                if (end - i >= 22) {
                    out.append(placeholder("drive", guidless.substring(i, end), "b!drive-%04d"));
                    i = end;
                    continue;
                }
            }
            if (boundary && wordChar(c)) {
                int end = run(guidless, i, ch -> wordChar(ch) || ch == '@');
                out.append(word(guidless.substring(i, end)));
                i = end;
                continue;
            }
            out.append(c);
            i++;
        }
        return out.toString();
    }

    private String word(String word) {
        if (emailLike(word)) {
            return placeholder("email", word.toLowerCase(Locale.ROOT), "email-%d.example.invalid");
        }
        String lower = word.toLowerCase(Locale.ROOT);
        for (String suffix : TENANT_HOSTS) {
            if (lower.endsWith(suffix) && lower.length() > suffix.length()) {
                return placeholder("host", lower, "tenant-%d.example");
            }
        }
        if (word.length() >= 30 && word.chars().allMatch(
                ch -> (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9'))) {
            return placeholder("item", word, "item-%04d");
        }
        return word;
    }

    private String guids(String input) {
        StringBuilder out = new StringBuilder(input.length());
        int i = 0;
        while (i < input.length()) {
            if (guidAt(input, i)) {
                String guid = input.substring(i, i + 36).toLowerCase(Locale.ROOT);
                out.append(guid.startsWith(GUID_PREFIX) ? guid
                        : placeholder("guid", guid, GUID_PREFIX + "%012d"));
                i += 36;
            } else {
                out.append(input.charAt(i++));
            }
        }
        return out.toString();
    }

    private String placeholder(String kind, String real, String format) {
        return placeholders.computeIfAbsent(kind + "\n" + real,
                key -> String.format(Locale.ROOT, format, counters.merge(kind, 1, Integer::sum)));
    }

    /** Whether a GUID (8-4-4-4-12 hex digits) starts at {@code at}. */
    static boolean guidAt(String text, int at) {
        if (at + 36 > text.length()) {
            return false;
        }
        for (int k = 0; k < 36; k++) {
            char c = text.charAt(at + k);
            boolean dash = k == 8 || k == 13 || k == 18 || k == 23;
            if (dash ? c != '-' : !hex(c)) {
                return false;
            }
        }
        return true;
    }

    /**
     * Whether a word is an email address or UPN. OData annotations such as
     * {@code fields@odata.context} are not.
     */
    static boolean emailLike(String word) {
        int at = word.indexOf('@');
        if (at <= 0 || at != word.lastIndexOf('@')) {
            return false;
        }
        String domain = word.substring(at + 1).toLowerCase(Locale.ROOT);
        if (domain.startsWith("odata.") || domain.startsWith("microsoft.graph.")) {
            return false;
        }
        int dot = domain.indexOf('.');
        return dot > 0 && dot < domain.length() - 1;
    }

    /** End of the URL that starts at {@code at}; {@code at} when none does. */
    static int urlEnd(String text, int at) {
        if (!text.startsWith("https://", at) && !text.startsWith("http://", at)) {
            return at;
        }
        return run(text, at, c -> !Character.isWhitespace(c) && c != '"' && c != '<' && c != '>');
    }

    /** {@code scheme://host[:port]} of a URL. */
    static String origin(String url) {
        int start = url.indexOf("://");
        if (start < 0) {
            return "";
        }
        int end = start + 3;
        while (end < url.length() && "/?#".indexOf(url.charAt(end)) < 0) {
            end++;
        }
        return url.substring(0, end);
    }

    /** One character test. */
    @FunctionalInterface
    interface CharTest {
        boolean test(char c);
    }

    static int run(String text, int from, CharTest test) {
        int end = from;
        while (end < text.length() && test.test(text.charAt(end))) {
            end++;
        }
        return end;
    }

    static boolean wordChar(char c) {
        return alnum(c) || c == '.' || c == '_' || c == '-';
    }

    private static boolean idChar(char c) {
        return alnum(c) || c == '_' || c == '-';
    }

    private static boolean tokenChar(char c) {
        return idChar(c) || c == '.';
    }

    private static boolean alnum(char c) {
        return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9');
    }

    private static boolean hex(char c) {
        return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
    }

    /**
     * What in {@code text} still looks like tenant data: a token, a tenant host,
     * an email address, or a GUID that is not a placeholder.
     */
    static List<String> violations(String text) {
        List<String> found = new ArrayList<>();
        String lower = text.toLowerCase(Locale.ROOT);
        for (String suffix : TENANT_HOSTS) {
            if (lower.contains(suffix.substring(1))) {
                found.add("tenant host (" + suffix.substring(1) + ")");
            }
        }
        for (int i = 0; i < text.length(); i++) {
            if (guidAt(text, i)) {
                if (!lower.startsWith(GUID_PREFIX, i)) {
                    found.add("GUID at offset " + i);
                }
                i += 35;
            } else if (text.startsWith("eyJ", i)) {
                int end = run(text, i, RecordingScrubber::idChar);
                if (end - i >= 10 && text.startsWith(".eyJ", end)) {
                    found.add("token at offset " + i);
                }
            } else if (text.charAt(i) == '@') {
                int start = i;
                while (start > 0 && wordChar(text.charAt(start - 1))) {
                    start--;
                }
                int end = run(text, i + 1, RecordingScrubber::wordChar);
                if (emailLike(text.substring(start, end))) {
                    found.add("email address at offset " + start);
                }
            }
        }
        return found;
    }
}
