package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import com.fasterxml.jackson.databind.JsonNode;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.net.http.HttpClient;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/** Records Graph calls made against the fake server, then replays them without it. */
class RecordingRoundTripTest {

    private static final String DRIVE = "b!Zm9vYmFyYmF6cXV4LWFiY2RlZl9naGlqa2xtbm9w";
    private static final String ITEM = "01ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
    private static final String USER = "3f2504e0-4f89-11d3-9a0c-0305e82c3301";

    private FakeGraphServer fake;
    private final Map<String, byte[]> blobs = new ConcurrentHashMap<>();
    private final byte[] content = "the file content\n".getBytes(StandardCharsets.UTF_8);

    @BeforeEach
    void start() throws IOException {
        fake = FakeGraphServer.start();
        blobs.put(Recordings.sha256(content), content);
    }

    @AfterEach
    void stop() {
        fake.close();
    }

    /** The calls both the recording and the replaying client make. */
    private static String calls(GraphClient graph, String driveId, byte[] upload) throws Exception {
        JsonNode created = graph.putBytes("/drives/" + driveId + "/root:/a.txt:/content", upload,
                upload.length, "text/plain", Map.of("if-match", "\"e1\""));
        String itemId = created.path("id").asText();
        JsonNode patched = graph.execute(new GraphClient.Call("PATCH",
                "/drives/" + driveId + "/items/" + itemId, Map.of(),
                GraphClient.object().put("name", "b.txt"))).body();
        byte[] downloaded = graph.getBytes("/drives/" + driveId + "/items/" + itemId + "/content");
        return itemId + " " + patched.path("createdBy").path("user").path("displayName").asText()
                + " " + new String(downloaded, StandardCharsets.UTF_8);
    }

    @Test
    void aRecordedRunReplaysWithoutTheServerAndRefusesOtherRequests(@TempDir Path directory)
            throws Exception {
        String base = fake.baseUrl();
        String itemJson = "{\"id\":\"" + ITEM + "\",\"eTag\":\"\\\"{" + USER + "},1\\\"\","
                + "\"createdBy\":{\"user\":{\"id\":\"" + USER + "\",\"displayName\":\"Ada\","
                + "\"email\":\"ada@contoso.com\"}}}";
        fake.stub("PUT", "/drives/" + DRIVE + "/root:/a.txt:/content",
                FakeGraphServer.Stub.json(itemJson));
        fake.stub("PATCH", "/drives/" + DRIVE + "/items/" + ITEM,
                FakeGraphServer.Stub.json(itemJson));
        fake.stub("/drives/" + DRIVE + "/items/" + ITEM + "/content", FakeGraphServer.empty(302,
                Map.of("Location", "http://localhost:" + base.substring(base.lastIndexOf(':') + 1)
                        + "/download/1?tempauth=eyJ0eXAiOiJKV1QifQ.eyJhdWQiOiJ4In0.c2ln")));
        fake.stub("/download/1", FakeGraphServer.Stub.bytes(content, "text/plain"));

        RecordingScrubber scrubber = new RecordingScrubber(base);
        RecordingTransport recorder = new RecordingTransport(
                HttpClient.newHttpClient()::send, scrubber, blobs, base);
        String live = calls(new GraphClient(base, () -> "secret-token",
                AuthMode.AUTH_MODE_APPLICATION, recorder), DRIVE, content);
        assertThat(live).isEqualTo(ITEM + " Ada the file content\n");
        assertThat(recorder.save(directory)).isEqualTo(4);
        fake.close();

        List<Path> files = Recordings.exchanges(directory);
        assertThat(files).extracting(file -> file.getFileName().toString()).containsExactly(
                "001-PUT-drives-root-a.txt-content.json", "002-PATCH-drives-items.json",
                "003-GET-drives-items-content.json", "004-GET-external.json");
        for (Path file : files) {
            String text = Files.readString(file);
            assertThat(text).doesNotContain("secret-token", ITEM, DRIVE, USER, "Ada", "contoso",
                    "tempauth");
        }
        JsonNode put = Recordings.JSON.readTree(files.get(0).toFile());
        assertThat(put.path("request").path("url").asText())
                .isEqualTo(base + "/drives/b!drive-0001/root:/a.txt:/content");
        assertThat(put.path("request").path("authorized").asBoolean()).isTrue();
        assertThat(put.path("request").path("headers").path("if-match").asText())
                .isEqualTo("\"e1\"");
        assertThat(put.path("request").path("body").path("sha256").asText())
                .isEqualTo(Recordings.sha256(content));
        JsonNode download = Recordings.JSON.readTree(files.get(3).toFile());
        assertThat(download.path("request").path("authorized").asBoolean()).isFalse();
        assertThat(download.path("response").path("body").path("blob").asText())
                .isEqualTo(Recordings.sha256(content));

        ReplayTransport replay = new ReplayTransport(blobs);
        GraphClient replayed = new GraphClient(base, () -> "another-token",
                AuthMode.AUTH_MODE_APPLICATION, replay);
        replay.begin(directory);
        assertThat(calls(replayed, "b!drive-0001", content))
                .isEqualTo("item-0001 Identity 1 the file content\n");
        assertThat(replay.finish()).isEmpty();

        replay.begin(directory);
        assertThatThrownBy(() -> calls(replayed, "b!drive-0001",
                "other bytes".getBytes(StandardCharsets.UTF_8)))
                .isInstanceOf(IOException.class).hasMessageContaining("replay mismatch");
        assertThat(replay.finish()).hasSize(2);
    }
}
