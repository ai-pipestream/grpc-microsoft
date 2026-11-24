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
    /** Domain of every placeholder email address. */
    static final String EMAIL_DOMAIN = "example.invalid";
    /** Start of every placeholder URL. */
    static final String URL_PREFIX = "https://redacted.example/url/";
    /** A reserved documentation origin; URLs on it are kept. */
    static final String DOCUMENTATION_ORIGIN = "https://example.com";
    /** Origin of the XML schema namespaces, which name a standard and no tenant. */
    static final String XML_SCHEMA_ORIGIN = "http://www.w3.org";
    /** A percent-encoded {@code @}. */
    static final String ENCODED_AT = "%40";

    /** Start of a site permission id that encodes the application and tenant it was granted to. */
    static final String ENCODED_GRANTEE = "aTowaS50fG1zLnNwLmV4dH";
    /** Shortest permission id that is replaced wherever it appears, not only as an {@code id}. */
    private static final int ENCODED_ID_LENGTH = 16;

    /** JSON objects whose {@code displayName} names a person, group or app. */
    private static final Set<String> IDENTITIES = Set.of("user", "application", "device", "group",
            "siteUser", "siteGroup", "sharePointGroup");
    /** {@code @odata.type} values of directory objects. */
    private static final Set<String> DIRECTORY_TYPES = Set.of("#microsoft.graph.user",
            "#microsoft.graph.group", "#microsoft.graph.device", "#microsoft.graph.servicePrincipal",
            "#microsoft.graph.orgContact");
    /** Properties of a directory object that say nothing about who it is. */
    private static final Set<String> DIRECTORY_KEPT = Set.of("id", "@odata.type", "userType",
            "visibility", "groupTypes", "preferredLanguage");
    /** Properties of a SharePoint principal that hold its name. */
    private static final Set<String> PRINCIPAL_NAMES = Set.of("Title", "Description", "OwnerTitle");
    /** Properties of a SharePoint change log entry that name no person, place or title. */
    private static final Set<String> CHANGE_KEPT = Set.of("odata.type", "Time");
    /** Query parameters whose value is a paging or delta token. */
    private static final List<String> CURSOR_KEYS = List.of("deltatoken=", "skiptoken=");
    /** Host suffixes that name the tenant. */
    private static final List<String> TENANT_HOSTS = List.of(".sharepoint.com",
            ".onmicrosoft.com", ".sharepoint.us", ".sharepoint.de", ".sharepoint.cn");

    private final String graphOrigin;
    private final Map<String, String> placeholders = new HashMap<>();
    private final Map<String, Integer> counters = new HashMap<>();
    /** Real permission ids long enough to be an encoded grantee, to their placeholders. */
    private final Map<String, String> permissionIds = new HashMap<>();
    /** Whether the body being scrubbed is a Graph listing of users, groups or directory objects. */
    private boolean directoryListing;

    /**
     * Creates a scrubber.
     *
     * @param graphBaseUrl the Graph base URL; URLs on its origin keep their path
     */
    RecordingScrubber(String graphBaseUrl) {
        this.graphOrigin = origin(graphBaseUrl);
    }

    /** Placeholder of the address named with {@link #joinProof}. */
    static final String JOIN_PROOF = "join-proof@" + EMAIL_DOMAIN;

    /** Names the one address that always gets {@link #JOIN_PROOF}, in any case. */
    synchronized void joinProof(String email) {
        if (email != null && !email.isBlank()) {
            placeholders.put("email\n" + email.strip().toLowerCase(Locale.ROOT), JOIN_PROOF);
        }
    }

    /** Scrubs every string value of a JSON tree; keys stay. */
    synchronized JsonNode json(JsonNode node) {
        String context = node.path("@odata.context").asText("");
        directoryListing = context.contains("$metadata#users") || context.contains("$metadata#groups")
                || context.contains("$metadata#directoryObjects");
        try {
            return json(node, "");
        } finally {
            directoryListing = false;
        }
    }

    private JsonNode json(JsonNode node, String parentKey) {
        if (node.isObject()) {
            ObjectNode out = JsonNodeFactory.instance.objectNode();
            // A delta of users or groups can hold only an id and the changed properties.
            boolean directory = directoryObject(node)
                    || (directoryListing && parentKey.equals("value") && node.has("id"));
            boolean principal = node.has("LoginName");
            boolean change = node.has("ChangeToken") && node.has("ChangeType");
            for (Map.Entry<String, JsonNode> entry : node.properties()) {
                JsonNode value = entry.getValue();
                if (directory && !DIRECTORY_KEPT.contains(entry.getKey())
                        && !entry.getKey().endsWith("DateTime") && (value.isTextual()
                        || (value.isArray() && !value.isEmpty() && value.get(0).isTextual()))) {
                    out.set(entry.getKey(), directoryValue(value));
                } else if (change && !CHANGE_KEPT.contains(entry.getKey())) {
                    out.set(entry.getKey(), changeValue(value));
                } else if (principal && value.isTextual() && (PRINCIPAL_NAMES.contains(entry.getKey())
                        || (entry.getKey().equals("LoginName") && value.asText().indexOf('|') < 0))) {
                    out.put(entry.getKey(), name("principal", value.asText(), "Principal %d"));
                } else if (value.isTextual() && (node.has(entry.getKey() + "LookupId")
                        || (entry.getKey().equals("LookupValue") && node.has("LookupId")))) {
                    // The shown value of a lookup or person column can be a person's name.
                    out.put(entry.getKey(), name("lookup", value.asText(), "Lookup value %d"));
                } else if (entry.getKey().equals("NameId") && value.isTextual()) {
                    out.put(entry.getKey(), name("nameid", value.asText(), "nameid-%04d"));
                } else if (entry.getKey().equals("displayName") && value.isTextual()
                        && IDENTITIES.contains(parentKey)) {
                    out.put(entry.getKey(), placeholder("name", value.asText(), "Identity %d"));
                } else if (entry.getKey().equals("id") && value.isTextual() && permission(node)) {
                    // A permission id can be the grantee's sign-in name, base64-encoded.
                    String id = placeholder("permission", value.asText(), "permission-%04d");
                    if (value.asText().length() >= ENCODED_ID_LENGTH) {
                        permissionIds.put(value.asText(), id);
                    }
                    out.put(entry.getKey(), id);
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

    /** Whether a JSON object is a permission; a listing of site permissions leaves the roles out. */
    private static boolean permission(JsonNode node) {
        return node.path("roles").isArray() || node.has("grantedToIdentitiesV2")
                || node.has("grantedToIdentities") || node.has("grantedToV2");
    }

    /** Whether a JSON object is an Entra user, group or other directory object. */
    private static boolean directoryObject(JsonNode node) {
        return node.has("userPrincipalName") || node.has("mailNickname") || node.has("securityEnabled")
                || DIRECTORY_TYPES.contains(node.path("@odata.type").asText());
    }

    /** A directory property: addresses become placeholder addresses, anything else a numbered value. */
    private JsonNode directoryValue(JsonNode value) {
        if (value.isArray()) {
            ArrayNode out = JsonNodeFactory.instance.arrayNode();
            value.forEach(element -> out.add(element.isTextual() ? directoryValue(element) : element));
            return out;
        }
        String text = value.asText();
        String scrubbed = text(text);
        return JsonNodeFactory.instance.textNode(scrubbed.equals(text)
                ? name("directory", text, "Directory value %d") : scrubbed);
    }

    /**
     * A property of a change log entry: editors, titles and paths become numbered values; ids are
     * GUIDs and keep their placeholders.
     */
    private JsonNode changeValue(JsonNode value) {
        if (value.isObject()) {
            ObjectNode out = JsonNodeFactory.instance.objectNode();
            value.properties().forEach(entry -> out.set(entry.getKey(), changeValue(entry.getValue())));
            return out;
        }
        if (value.isArray()) {
            ArrayNode out = JsonNodeFactory.instance.arrayNode();
            value.forEach(element -> out.add(changeValue(element)));
            return out;
        }
        if (!value.isTextual()) {
            return value;
        }
        String scrubbed = text(value.asText());
        return JsonNodeFactory.instance.textNode(scrubbed.equals(value.asText())
                ? name("change", scrubbed, "Change value %d") : scrubbed);
    }

    private String name(String kind, String real, String format) {
        return real.isEmpty() ? real : placeholder(kind, real, format);
    }

    /** Scrubs free text: URLs, tokens, GUIDs, drive and item ids, hosts and emails. */
    synchronized String text(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }
        String known = input;
        for (Map.Entry<String, String> id : permissionIds.entrySet()) {
            // A permission id seen in an answer, now in the URL of a later request.
            known = known.replace(id.getKey(), id.getValue());
        }
        String guidless = guids(known);
        StringBuilder out = new StringBuilder(guidless.length());
        int i = 0;
        while (i < guidless.length()) {
            int url = urlEnd(guidless, i);
            if (url > i) {
                String value = guidless.substring(i, url);
                String origin = origin(value);
                int api = apiStart(value);
                if (origin.equalsIgnoreCase(XML_SCHEMA_ORIGIN)) {
                    out.append(value);
                    i = url;
                } else if (origin.equalsIgnoreCase(DOCUMENTATION_ORIGIN)) {
                    // A reserved documentation address, which the tests write themselves.
                    out.append(origin);
                    i += origin.length();
                } else if (origin.equalsIgnoreCase(graphOrigin)) {
                    out.append(graphOrigin);
                    i += origin.length();
                } else if (api > 0) {
                    // A site's REST API: the site's address is replaced, the API path stays.
                    out.append(placeholder("url", value.substring(0, api), URL_PREFIX + "%d"));
                    i += api;
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
            if (boundary && guidless.startsWith("b%21", i)) {
                // The same drive id inside a URL, with its '!' percent-encoded.
                int end = run(guidless, i + 4, RecordingScrubber::idChar);
                if (end - i >= 24) {
                    out.append(placeholder("drive", "b!" + guidless.substring(i + 4, end),
                            "b!drive-%04d").replace("b!", "b%21"));
                    i = end;
                    continue;
                }
            }
            int cursor = cursorValue(guidless, i);
            if (cursor > i) {
                // The paging or delta token of a Graph link; it encodes tenant state.
                int end = run(guidless, cursor, ch -> !Character.isWhitespace(ch) && "&\"'<>".indexOf(ch) < 0);
                String token = guidless.substring(cursor, end);
                out.append(guidless, i, cursor).append(token.isEmpty() || token.equals("latest")
                        || token.startsWith("cursor-") ? token : placeholder("cursor", token, "cursor-%04d"));
                i = end;
                continue;
            }
            if (c == '%' && !guidless.startsWith(ENCODED_AT, i) && i + 2 < guidless.length()
                    && hex(guidless.charAt(i + 1)) && hex(guidless.charAt(i + 2))) {
                // A percent-encoded character of a URL; the word after it starts fresh.
                out.append(guidless, i, i + 3);
                i += 3;
                continue;
            }
            if ((boundary || escaped(guidless, i)) && wordChar(c)) {
                int end = i;
                while (end < guidless.length()) {
                    char ch = guidless.charAt(end);
                    if (wordChar(ch) || ch == '@') {
                        end++;
                    } else if (guidless.startsWith(ENCODED_AT, end)) {
                        end += ENCODED_AT.length();
                    } else {
                        break;
                    }
                }
                String word = guidless.substring(i, end);
                // An address inside a URL has its '@' percent-encoded.
                out.append(word.contains(ENCODED_AT) && emailLike(word.replace(ENCODED_AT, "@"))
                        ? word(word.replace(ENCODED_AT, "@")).replace("@", ENCODED_AT)
                        : word(word));
                i = end;
                continue;
            }
            out.append(c);
            i++;
        }
        return out.toString();
    }

    /** Whether the character at {@code at} follows a percent-encoded character. */
    private static boolean escaped(String text, int at) {
        return at >= 3 && text.charAt(at - 3) == '%' && hex(text.charAt(at - 2))
                && hex(text.charAt(at - 1)) && !text.startsWith(ENCODED_AT, at - 3);
    }

    private String word(String word) {
        if (emailLike(word)) {
            return placeholder("email", word.toLowerCase(Locale.ROOT), "email-%d@" + EMAIL_DOMAIN);
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

    /** Where {@code /_api/web} or {@code /_api/site} starts in the URL of a site's REST API; -1 in any other URL. */
    static int apiStart(String url) {
        for (String root : List.of("/_api/web", "/_api/site")) {
            int at = url.indexOf(root);
            int end = at + root.length();
            if (at > 0 && (end == url.length() || "/?".indexOf(url.charAt(end)) >= 0)) {
                return at;
            }
        }
        return -1;
    }

    /** Where the value of a {@code deltatoken=} or {@code skiptoken=} parameter starts; {@code at} when none starts here. */
    static int cursorValue(String text, int at) {
        for (String key : CURSOR_KEYS) {
            if (text.regionMatches(true, at, key, 0, key.length())) {
                return at + key.length();
            }
        }
        return at;
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
     * an email address or a GUID that is not a placeholder.
     */
    static List<String> violations(String text) {
        List<String> found = new ArrayList<>();
        if (text.contains(ENCODED_AT)) {
            // Addresses inside URLs, with the '@' percent-encoded and an encoded character before.
            if (violations(text.replace(ENCODED_AT, "@").replace('%', ' ')).stream()
                    .anyMatch(finding -> finding.startsWith("email address"))) {
                found.add("percent-encoded email address");
            }
        }
        String lower = text.toLowerCase(Locale.ROOT);
        for (String suffix : TENANT_HOSTS) {
            if (lower.contains(suffix.substring(1))) {
                found.add("tenant host (" + suffix.substring(1) + ")");
            }
        }
        if (text.contains(ENCODED_GRANTEE)) {
            found.add("encoded grantee of a site permission");
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
                if (emailLike(text.substring(start, end))
                        && !text.substring(i + 1, end).equalsIgnoreCase(EMAIL_DOMAIN)) {
                    found.add("email address at offset " + start);
                }
            }
        }
        return found;
    }
}
