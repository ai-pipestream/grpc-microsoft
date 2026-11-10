package ai.pipestream.microsoft;

import com.fasterxml.jackson.databind.JsonNode;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class RecordingScrubberTest {

    private static final String GRAPH = "https://graph.microsoft.com/v1.0";
    private static final String GUID_A = "3f2504e0-4f89-11d3-9a0c-0305e82c3301";
    private static final String GUID_B = "9B2C1D3E-0000-4AAA-8BBB-123456789ABC";
    private static final String ITEM = "01ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
    private static final String DRIVE = "b!Zm9vYmFyYmF6cXV4LWFiY2RlZl9naGlqa2xtbm9w";

    private final RecordingScrubber scrubber = new RecordingScrubber(GRAPH);

    @Test
    void theSameRealValueAlwaysBecomesTheSamePlaceholder() {
        String first = scrubber.text("contoso.sharepoint.com," + GUID_A + "," + GUID_B);
        assertThat(first).isEqualTo("tenant-1.example," + RecordingScrubber.GUID_PREFIX
                + "000000000001," + RecordingScrubber.GUID_PREFIX + "000000000002");
        assertThat(scrubber.text("\"{" + GUID_B.toLowerCase() + "},4\""))
                .isEqualTo("\"{" + RecordingScrubber.GUID_PREFIX + "000000000002},4\"");
        assertThat(scrubber.text(first)).isEqualTo(first);
    }

    @Test
    void graphUrlsKeepTheirRouteAndOtherUrlsAreReplacedWhole() {
        assertThat(scrubber.text(GRAPH + "/drives/" + DRIVE + "/items/" + ITEM + ":/a.txt:/content"))
                .isEqualTo(GRAPH + "/drives/b!drive-0001/items/item-0001:/a.txt:/content");
        String upload = "https://contoso.sharepoint.com/_api/v2.0/up?guid='" + GUID_A
                + "'&tempauth=eyJ0eXAiOiJKV1QifQ.eyJhdWQiOiJ4In0.c2ln";
        assertThat(scrubber.text(upload)).isEqualTo(RecordingScrubber.URL_PREFIX + "1");
        assertThat(scrubber.text("see " + upload + " now"))
                .isEqualTo("see " + RecordingScrubber.URL_PREFIX + "1 now");
        assertThat(scrubber.text("https://contoso-my.sharepoint.com/other"))
                .isEqualTo(RecordingScrubber.URL_PREFIX + "2");
    }

    @Test
    void tokensEmailsAndHostsAreReplaced() {
        assertThat(scrubber.text("Bearer eyJ0eXAiOiJKV1QifQ.eyJhdWQiOiJ4In0.c2lnbmF0dXJl"))
                .isEqualTo("Bearer token-redacted");
        assertThat(scrubber.text("i:0#.f|membership|Ada.Lovelace@contoso.onmicrosoft.com"))
                .isEqualTo("i:0#.f|membership|email-1.example.invalid");
        assertThat(scrubber.text("ada.lovelace@contoso.onmicrosoft.com"))
                .isEqualTo("email-1.example.invalid");
        assertThat(scrubber.text("host contoso.sharepoint.com")).isEqualTo("host tenant-1.example");
        assertThat(scrubber.text("pipestream-live-write-20260102-030405-1a2b3c4d"))
                .isEqualTo("pipestream-live-write-20260102-030405-1a2b3c4d");
    }

    @Test
    void jsonValuesAreScrubbedAndKeysAndOwnNamesStay() throws Exception {
        JsonNode scrubbed = scrubber.json(Recordings.JSON.readTree("""
                {"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#drives",
                 "fields@odata.context":"x",
                 "id":"%s","displayName":"pipestream-live-write-1",
                 "webUrl":"https://contoso.sharepoint.com/sites/Team/Docs",
                 "createdBy":{"user":{"id":"%s","displayName":"Ada Lovelace",
                                      "email":"ada@contoso.com"},
                              "application":{"displayName":"Crawler app"}},
                 "lastModifiedBy":{"user":{"displayName":"Ada Lovelace"}},
                 "size":12,"value":[{"id":"%s"}]}
                """.formatted(ITEM, GUID_A, ITEM)));

        assertThat(scrubbed.path("@odata.context").asText())
                .isEqualTo("https://graph.microsoft.com/v1.0/$metadata#drives");
        assertThat(scrubbed.has("fields@odata.context")).isTrue();
        assertThat(scrubbed.path("id").asText()).isEqualTo("item-0001");
        assertThat(scrubbed.path("value").path(0).path("id").asText()).isEqualTo("item-0001");
        assertThat(scrubbed.path("displayName").asText()).isEqualTo("pipestream-live-write-1");
        assertThat(scrubbed.path("webUrl").asText()).startsWith(RecordingScrubber.URL_PREFIX);
        JsonNode user = scrubbed.path("createdBy").path("user");
        assertThat(user.path("displayName").asText()).isEqualTo("Identity 1");
        assertThat(user.path("email").asText()).isEqualTo("email-1.example.invalid");
        assertThat(user.path("id").asText()).startsWith(RecordingScrubber.GUID_PREFIX);
        assertThat(scrubbed.path("createdBy").path("application").path("displayName").asText())
                .isEqualTo("Identity 2");
        assertThat(scrubbed.path("lastModifiedBy").path("user").path("displayName").asText())
                .isEqualTo("Identity 1");
        assertThat(scrubbed.path("size").asInt()).isEqualTo(12);
        assertThat(RecordingScrubber.violations(scrubbed.toString())).isEmpty();
    }

    @Test
    void violationsNameWhatStillLooksLikeTenantData() {
        assertThat(RecordingScrubber.violations("{\"fields@odata.context\":\"x\",\"id\":\""
                + RecordingScrubber.GUID_PREFIX + "000000000007\",\"@odata.etag\":\"1\"}"))
                .isEmpty();
        assertThat(RecordingScrubber.violations("id " + GUID_A)).hasSize(1);
        assertThat(RecordingScrubber.violations("ada@contoso.com")).hasSize(1);
        assertThat(RecordingScrubber.violations("x CONTOSO.SharePoint.com y")).hasSize(1);
        assertThat(RecordingScrubber.violations("eyJ0eXAiOiJKV1QifQ.eyJhdWQiOiJ4In0.c2ln"))
                .hasSize(1);
    }
}
