package ai.pipestream.microsoft;

import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.assertj.core.api.Assertions.assertThat;

/** Fails when a committed recording holds a token, a tenant host, an email or a real GUID. */
class RecordingsHygieneTest {

    @Test
    void noRecordingHoldsTenantData() throws Exception {
        Path root = Recordings.dir();
        List<String> found = new ArrayList<>();
        if (Files.isDirectory(root)) {
            try (Stream<Path> tree = Files.walk(root)) {
                for (Path file : tree.filter(Files::isRegularFile).sorted().toList()) {
                    String text = Files.readString(file, StandardCharsets.UTF_8);
                    RecordingScrubber.violations(text)
                            .forEach(violation -> found.add(root.relativize(file) + ": " + violation));
                }
            }
        }
        assertThat(found).isEmpty();
    }
}
