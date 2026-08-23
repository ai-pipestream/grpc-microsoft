package ai.pipestream.microsoft;

import ai.pipestream.okf.OkfBundle;
import ai.pipestream.okf.OkfOutput;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;
import java.util.Objects;

/**
 * Uploads an OKF tree, zip, and sibling WARC to a SharePoint drive folder.
 * Large files use a Graph upload session (no {@code Authorization} on the
 * session URL).
 */
public final class SharePointOkfPublisher {

    /** Destination drive id. */
    public static final String ENV_DRIVE_ID = "OKF_SPO_DRIVE_ID";
    /** Destination folder path (drive-root relative). */
    public static final String ENV_FOLDER_PATH = "OKF_SPO_FOLDER_PATH";

    private final GraphFiles files;
    private final String driveId;
    private final String folderPath;

    /**
     * Creates a publisher.
     *
     * @param files Graph files API
     * @param driveId destination drive
     * @param folderPath folder from drive root; blank is the root
     */
    public SharePointOkfPublisher(GraphFiles files, String driveId, String folderPath) {
        this.files = Objects.requireNonNull(files, "files");
        this.driveId = Objects.requireNonNull(driveId, "driveId");
        this.folderPath = folderPath == null || folderPath.isBlank() ? "" : folderPath;
    }

    /**
     * Whether SharePoint upload is configured.
     *
     * @return true when {@link #ENV_DRIVE_ID} is set
     */
    public static boolean enabled() {
        String drive = System.getenv(ENV_DRIVE_ID);
        return drive != null && !drive.isBlank();
    }

    /**
     * Builds a publisher from the process environment.
     *
     * @param files Graph files API
     * @return the publisher
     */
    public static SharePointOkfPublisher fromEnvironment(GraphFiles files) {
        return new SharePointOkfPublisher(files, System.getenv(ENV_DRIVE_ID),
                System.getenv(ENV_FOLDER_PATH));
    }

    /**
     * Uploads every OKF file plus optional zip and warc siblings.
     *
     * @param bundle the OKF files
     * @param output local destinations already written (zip/warc paths may be null)
     * @throws IOException if an upload fails
     * @throws InterruptedException if a call is interrupted
     */
    public void publish(OkfBundle bundle, OkfOutput output)
            throws IOException, InterruptedException {
        for (var entry : bundle.files().entrySet()) {
            upload(entry.getKey(), entry.getValue());
        }
        if (output.zip() != null && Files.isRegularFile(output.zip())) {
            upload(output.zip().getFileName().toString(), Files.readAllBytes(output.zip()));
        }
        if (output.warc() != null && Files.isRegularFile(output.warc())) {
            upload(output.warc().getFileName().toString(), Files.readAllBytes(output.warc()));
        }
    }

    private void upload(String relativePath, byte[] content)
            throws IOException, InterruptedException {
        String dest = join(folderPath, relativePath);
        files.uploadOrSession(driveId, dest, content, mime(relativePath));
    }

    static String mime(String path) {
        String lower = path.toLowerCase(Locale.ROOT);
        if (lower.endsWith(".md")) {
            return "text/markdown; charset=utf-8";
        }
        if (lower.endsWith(".html")) {
            return "text/html; charset=utf-8";
        }
        if (lower.endsWith(".zip")) {
            return "application/zip";
        }
        if (lower.endsWith(".warc.gz") || lower.endsWith(".gz")) {
            return "application/gzip";
        }
        return "application/octet-stream";
    }

    static String join(String folder, String relative) {
        String left = folder == null ? "" : folder.replace('\\', '/');
        String right = relative == null ? "" : relative.replace('\\', '/');
        while (left.endsWith("/")) {
            left = left.substring(0, left.length() - 1);
        }
        while (right.startsWith("/")) {
            right = right.substring(1);
        }
        if (left.isEmpty()) {
            return "/" + right;
        }
        if (right.isEmpty()) {
            return left.startsWith("/") ? left : "/" + left;
        }
        return (left.startsWith("/") ? left : "/" + left) + "/" + right;
    }
}
