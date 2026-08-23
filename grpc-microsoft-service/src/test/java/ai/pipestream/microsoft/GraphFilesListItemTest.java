package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class GraphFilesListItemTest {

    private FakeGraphServer fake;
    private GraphFiles files;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token"));
    }

    @AfterEach
    void stop() {
        fake.close();
    }

    @Test
    void listItemFieldsOrEmptyReturnsFieldsObject() throws Exception {
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields",
                MicrosoftFixtures.listItemFieldsJson());
        ObjectNode fields = files.listItemFieldsOrEmpty("drive-1", "file-1");
        assertThat(fields.path("Title").asText()).isEqualTo("Notes");
        assertThat(fields.path("Count").asInt()).isEqualTo(3);
        assertThat(fields.has("@odata.etag")).isTrue();
    }

    @Test
    void listItemFieldsOrEmptySwallows404() throws Exception {
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields",
                new FakeGraphServer.Stub(404,
                        "{\"error\":{\"code\":\"itemNotFound\",\"message\":\"no list item\"}}"
                                .getBytes(StandardCharsets.UTF_8),
                        Map.of("content-type", "application/json")));
        assertThat(files.listItemFieldsOrEmpty("drive-1", "file-1")).isEmpty();
        assertThat(fake.requests()).noneMatch(r -> r.path().equals("/drives/drive-1/items/file-1")
                && r.query().isEmpty());
    }

    @Test
    void listItemFieldsOrEmptyRethrowsNon404() {
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields",
                new FakeGraphServer.Stub(500,
                        "{\"error\":{\"code\":\"fail\",\"message\":\"boom\"}}"
                                .getBytes(StandardCharsets.UTF_8),
                        Map.of("content-type", "application/json")));
        assertThatThrownBy(() -> files.listItemFieldsOrEmpty("drive-1", "file-1"))
                .isInstanceOf(GraphClient.GraphApiException.class)
                .satisfies(t -> assertThat(((GraphClient.GraphApiException) t).status()).isEqualTo(500));
    }

    @Test
    void listItemFieldsOrEmptyTreatsMissingFieldsAsEmpty() throws Exception {
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields", "{\"id\":\"li-1\"}");
        assertThat(files.listItemFieldsOrEmpty("drive-1", "file-1")).isEmpty();
    }

    @Test
    void listItemFieldsOnlyReturnsEmptyWhenItemExistsWithoutListItem() throws Exception {
        fake.stub("/drives/drive-1/items/file-1",
                MicrosoftFixtures.fileJson("file-1", "notes.txt", "drive-1"));
        fake.stub("/drives/drive-1/items/file-1/listItem?$expand=fields",
                new FakeGraphServer.Stub(404,
                        "{\"error\":{\"code\":\"itemNotFound\",\"message\":\"no list item\"}}"
                                .getBytes(StandardCharsets.UTF_8),
                        Map.of("content-type", "application/json")));
        assertThat(files.listItemFieldsOnly("drive-1", "file-1")).isEmpty();
        assertThat(fake.requests()).anyMatch(r -> "/drives/drive-1/items/file-1".equals(r.path()));
    }

    @Test
    void listItemFieldsOnlyRethrowsWhenItemIsAlsoMissing() {
        fake.stub("/drives/drive-1/items/missing/listItem?$expand=fields",
                new FakeGraphServer.Stub(404,
                        "{\"error\":{\"code\":\"itemNotFound\",\"message\":\"gone\"}}"
                                .getBytes(StandardCharsets.UTF_8),
                        Map.of("content-type", "application/json")));
        fake.stub("/drives/drive-1/items/missing",
                new FakeGraphServer.Stub(404,
                        "{\"error\":{\"code\":\"itemNotFound\",\"message\":\"gone\"}}"
                                .getBytes(StandardCharsets.UTF_8),
                        Map.of("content-type", "application/json")));
        assertThatThrownBy(() -> files.listItemFieldsOnly("drive-1", "missing"))
                .isInstanceOf(GraphClient.GraphApiException.class)
                .satisfies(t -> assertThat(((GraphClient.GraphApiException) t).status()).isEqualTo(404));
    }
}
