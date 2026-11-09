package ai.pipestream.microsoft;

import java.net.URI;
import java.net.URISyntaxException;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Validation for URLs the server is asked to fetch with its Graph token: the
 * {@code @odata.nextLink} / {@code @odata.deltaLink} / {@code Location} values
 * Graph returns, and above all the delta links a client hands back. A link is
 * followed only when it is on the configured Graph origin and under its base
 * path, and a client-supplied delta link must also name the delta route of the
 * drive, list, or site catalog the request is about. Anything else would let a
 * caller point the server's bearer token at a host of its choosing.
 */
public final class GraphLinks {

    private GraphLinks() {
    }

    /**
     * Removes every trailing {@code /} from {@code url}.
     *
     * @param url a URL or path
     * @return {@code url} without trailing slashes
     */
    public static String stripTrailingSlashes(String url) {
        int end = url.length();
        while (end > 0 && url.charAt(end - 1) == '/') {
            end--;
        }
        return url.substring(0, end);
    }

    /**
     * Whether {@code url} is an absolute URL on the same scheme, host and port as
     * {@code baseUrl} and under its path, with no user-info.
     *
     * @param baseUrl the configured Graph root, for example
     *        {@code https://graph.microsoft.com/v1.0}
     * @param url the URL to check
     * @return {@code true} when the server may send its token to {@code url}
     */
    public static boolean isUnderBase(String baseUrl, String url) {
        try {
            URI base = new URI(stripTrailingSlashes(baseUrl));
            URI target = new URI(url);
            if (!target.isAbsolute() || target.getRawUserInfo() != null
                    || target.getHost() == null || base.getHost() == null) {
                return false;
            }
            if (!base.getScheme().equalsIgnoreCase(target.getScheme())
                    || !base.getHost().equalsIgnoreCase(target.getHost())
                    || port(base) != port(target)) {
                return false;
            }
            String basePath = base.getRawPath() == null ? "" : base.getRawPath();
            String path = target.getRawPath() == null ? "" : target.getRawPath();
            return path.startsWith(basePath + "/");
        } catch (URISyntaxException e) {
            return false;
        }
    }

    /**
     * Requires {@code link} to be a delta link for {@code route} under
     * {@code baseUrl}: the path after the base must be the route segments
     * (compared case-insensitively) followed by exactly one segment that starts
     * with {@code delta} or {@code microsoft.graph.delta} (Graph returns both
     * {@code delta?token=...} and {@code delta(token='...')} forms).
     *
     * @param baseUrl the configured Graph root
     * @param link the client-supplied delta link
     * @param route the expected route segments before the delta function, for
     *        example {@code drives, {drive-id}, root}
     * @return {@code link}, unchanged
     * @throws IllegalArgumentException when the link is not on the Graph origin
     *         or not this resource's delta route
     */
    public static String requireDeltaLink(String baseUrl, String link, String... route) {
        if (link == null || link.isBlank()) {
            throw new IllegalArgumentException("delta link is empty");
        }
        if (!isUnderBase(baseUrl, link)) {
            throw new IllegalArgumentException("delta link is not on the configured Graph endpoint "
                    + stripTrailingSlashes(baseUrl));
        }
        List<String> segments = segmentsAfterBase(baseUrl, link);
        if (segments.size() != route.length + 1) {
            throw new IllegalArgumentException("delta link is not the delta route of "
                    + String.join("/", route));
        }
        for (int i = 0; i < route.length; i++) {
            if (!segments.get(i).equalsIgnoreCase(route[i])) {
                throw new IllegalArgumentException("delta link is not the delta route of "
                        + String.join("/", route));
            }
        }
        String function = segments.get(route.length).toLowerCase(Locale.ROOT);
        if (!function.startsWith("delta") && !function.startsWith("microsoft.graph.delta")) {
            throw new IllegalArgumentException("delta link does not call a delta function");
        }
        return link;
    }

    private static List<String> segmentsAfterBase(String baseUrl, String link) {
        String basePath = URI.create(stripTrailingSlashes(baseUrl)).getRawPath();
        String path = URI.create(link).getRawPath().substring(basePath == null ? 0 : basePath.length());
        List<String> out = new ArrayList<>();
        int start = 0;
        while (start <= path.length()) {
            int slash = path.indexOf('/', start);
            int end = slash < 0 ? path.length() : slash;
            String raw = path.substring(start, end);
            if (!raw.isEmpty()) {
                out.add(URLDecoder.decode(raw.replace("+", "%2B"), StandardCharsets.UTF_8));
            }
            if (slash < 0) {
                break;
            }
            start = slash + 1;
        }
        return out;
    }

    private static int port(URI uri) {
        if (uri.getPort() >= 0) {
            return uri.getPort();
        }
        return "https".equalsIgnoreCase(uri.getScheme()) ? 443 : 80;
    }
}
