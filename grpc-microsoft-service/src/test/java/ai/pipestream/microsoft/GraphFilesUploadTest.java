package ai.pipestream.microsoft;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.util.Arrays;

import static org.assertj.core.api.Assertions.assertThat;

class GraphFilesUploadTest {

    private FakeGraphServer fake;

    @AfterEach
    void stop() {
        if (fake != null) {
            fake.close();
        }
    }

    @Test
    void uploadSessionPutsChunksWithoutAuthorization() throws Exception {
        fake = FakeGraphServer.start();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        String uploadUrl = fake.baseUrl() + "/session-upload";
        fake.stub("/drives/drive-1/root:/big.bin:/createUploadSession",
                "{\"uploadUrl\":\"" + uploadUrl + "\"}");
        fake.stub("/session-upload", "{\"id\":\"big\",\"name\":\"big.bin\"}");

        byte[] content = new byte[GraphFiles.SIMPLE_UPLOAD_LIMIT + 1];
        Arrays.fill(content, (byte) 7);
        files.uploadOrSession("drive-1", "/big.bin", content, "application/octet-stream");

        assertThat(fake.requests()).anyMatch(r -> "POST".equals(r.method())
                && r.path().endsWith("/createUploadSession")
                && "Bearer token".equals(r.authorization()));
        assertThat(fake.requests()).filteredOn(r -> "/session-upload".equals(r.path()))
                .isNotEmpty()
                .allMatch(r -> r.authorization() == null)
                .allMatch(r -> r.contentRange() != null && r.contentRange().startsWith("bytes "));
    }

    @Test
    void simpleUploadStaysUnderLimit() throws Exception {
        fake = FakeGraphServer.start();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
        files.uploadOrSession("drive-1", "/notes.txt", "hi".getBytes(StandardCharsets.UTF_8),
                "text/plain");
        assertThat(fake.requests()).anyMatch(r -> "PUT".equals(r.method())
                && r.path().contains("notes.txt")
                && "Bearer token".equals(r.authorization()));
    }
}
