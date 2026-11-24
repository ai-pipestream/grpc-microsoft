package ai.pipestream.microsoft;

import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Set;

/** Where the change-feed recordings live and what a step saves besides Graph exchanges. */
final class ChangeRecordings {

    /** Default directory, relative to the module directory. */
    static final String DEFAULT_DIR = "src/test/resources/recordings/microsoft-changes";
    /** System property that overrides {@link #DEFAULT_DIR}. */
    static final String DIR_PROPERTY = "changeRecordings.dir";
    /** File of a step that holds the notification bodies Graph posted. */
    static final String NOTIFICATIONS_FILE = "notifications.json";
    /** Steps that wait for notifications and save them. */
    static final Set<String> NOTIFIED_STEPS = Set.of("09-notified-pass",
            "10-permission-notification", ChangeScenarios.LIFECYCLE_STEP);

    private ChangeRecordings() {
    }

    /** The recordings directory of this build. */
    static Path dir() {
        return Path.of(System.getProperty(DIR_PROPERTY, DEFAULT_DIR));
    }

    /** The file content a run with {@code suffix} uploads. */
    static byte[] content(String suffix) {
        return ("pipestream live change " + suffix).getBytes(StandardCharsets.UTF_8);
    }
}
