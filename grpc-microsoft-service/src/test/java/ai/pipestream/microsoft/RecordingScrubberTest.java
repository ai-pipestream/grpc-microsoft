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
                .isEqualTo("i:0#.f|membership|email-1@example.invalid");
        assertThat(scrubber.text("ada.lovelace@contoso.onmicrosoft.com"))
                .isEqualTo("email-1@example.invalid");
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
        assertThat(user.path("email").asText()).isEqualTo("email-1@example.invalid");
        assertThat(user.path("id").asText()).startsWith(RecordingScrubber.GUID_PREFIX);
        assertThat(scrubbed.path("createdBy").path("application").path("displayName").asText())
                .isEqualTo("Identity 2");
        assertThat(scrubbed.path("lastModifiedBy").path("user").path("displayName").asText())
                .isEqualTo("Identity 1");
        assertThat(scrubbed.path("size").asInt()).isEqualTo(12);
        assertThat(RecordingScrubber.violations(scrubbed.toString())).isEmpty();
    }

    @Test
    void aPercentEncodedDriveIdBecomesTheSamePlaceholder() {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        String id = "b!8FWe-5b5AUuTqSDN0boiUaukpeBeel5Co";
        assertThat(scrubber.text(id)).isEqualTo("b!drive-0001");
        assertThat(scrubber.text("https://graph.microsoft.com/v1.0/$metadata#drives('"
                + id.replace("b!", "b%21") + "')/root"))
                .isEqualTo("https://graph.microsoft.com/v1.0/$metadata#drives('b%21drive-0001')/root");
    }

    @Test
    void permissionIdsAreReplaced() throws Exception {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        JsonNode scrubbed = scrubber.json(Recordings.JSON.readTree("""
                {"value":[{"id":"aTowIy5mfG1lbWJlcnNoaXB8c29tZW9uZQ","roles":["write"]},
                          {"id":"7","name":"kept"}]}"""));
        assertThat(scrubbed.path("value").path(0).path("id").asText()).isEqualTo("permission-0001");
        assertThat(scrubbed.path("value").path(1).path("id").asText()).isEqualTo("7");
    }

    @Test
    void aShareIdIsReplacedAndTheLeakCheckSeesOne() throws Exception {
        String share = "u!aHR0cHM6Ly9jb250b3NvLnNoYXJlcG9pbnQuY29tLzp0Oi9zL1RlYW0vYWJj";
        JsonNode scrubbed = scrubber.json(Recordings.JSON.readTree("""
                {"value":[{"id":"%s","roles":["read"],"shareId":"%s","link":{"scope":"organization"}},
                          {"id":"aTowIy5mfG1lbWJlcnNoaXB8c29tZW9uZQ","roles":["write"],
                           "shareId":"aTowIy5mfG1lbWJlcnNoaXB8c29tZW9uZQ"}]}""".formatted(GUID_A, share)));

        assertThat(scrubbed.path("value").path(0).path("shareId").asText()).isEqualTo("share-0001");
        assertThat(scrubbed.path("value").path(1).path("shareId").asText())
                .isEqualTo(scrubbed.path("value").path(1).path("id").asText());
        assertThat(RecordingScrubber.violations("\"shareId\":\"" + share + "\""))
                .containsExactly("encoded sharing URL");
        assertThat(RecordingScrubber.violations(scrubbed.toString())).isEmpty();
    }

    @Test
    void theGroupOfASharingLinkOrOfLimitedAccessKeepsItsKind() throws Exception {
        JsonNode scrubbed = scrubber.json(Recordings.JSON.readTree("""
                {"value":[
                 {"Id":20,"LoginName":"SharingLinks.%1$s.OrganizationView.%2$s",
                  "Title":"SharingLinks.%1$s.OrganizationView.%2$s","Description":"made for a link"},
                 {"Id":21,"LoginName":"Limited Access System Group For List %1$s",
                  "Title":"Limited Access System Group For List %1$s"},
                 {"Id":22,"LoginName":"SharingLinks and more","Title":"SharingLinks and more"}],
                 "UniqueId":"%1$s"}""".formatted(GUID_A, GUID_B)));

        String item = scrubbed.path("UniqueId").asText();
        JsonNode link = scrubbed.path("value").path(0);
        assertThat(item).startsWith(RecordingScrubber.GUID_PREFIX);
        assertThat(link.path("LoginName").asText()).startsWith("SharingLinks." + item + ".OrganizationView.")
                .doesNotContain(GUID_B);
        assertThat(link.path("Title").asText()).isEqualTo(link.path("LoginName").asText());
        assertThat(link.path("Description").asText()).isEqualTo("Principal 1");
        assertThat(scrubbed.path("value").path(1).path("LoginName").asText())
                .isEqualTo("Limited Access System Group For List " + item);
        assertThat(scrubbed.path("value").path(2).path("LoginName").asText()).isEqualTo("Principal 2");
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

    @Test
    void aSiteRestUrlKeepsItsApiPathAndSharesThePlaceholderOfTheSite() {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        String site = scrubber.text("https://contoso.sharepoint.com/sites/Team");
        assertThat(site).isEqualTo(RecordingScrubber.URL_PREFIX + "1");
        assertThat(scrubber.text("https://contoso.sharepoint.com/sites/Team/_api/web/lists(guid'"
                + GUID_A + "')/roleassignments?$expand=Member"))
                .isEqualTo(site + "/_api/web/lists(guid'" + RecordingScrubber.GUID_PREFIX
                        + "000000000001')/roleassignments?$expand=Member");
        assertThat(scrubber.text("https://contoso.sharepoint.com/sites/Team/_api/site/getchanges"))
                .isEqualTo(site + "/_api/site/getchanges");
        assertThat(scrubber.text("https://contoso.sharepoint.com/sites/Team/_api/site?$select=CurrentChangeToken"))
                .isEqualTo(site + "/_api/site?$select=CurrentChangeToken");
        assertThat(scrubber.text("https://contoso.sharepoint.com/sites/Team/_api/sitepages/pages"))
                .isEqualTo(RecordingScrubber.URL_PREFIX + "2");
    }

    @Test
    void directoryObjectsAndSharePointPrincipalsLoseTheirNames() throws Exception {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        JsonNode user = scrubber.json(Recordings.JSON.readTree("""
                {"id":"u1","displayName":"Ada Lovelace","userPrincipalName":"ada@contoso.com",
                 "jobTitle":"Engineer","businessPhones":["+1 555 0100"],"userType":"Member",
                 "createdDateTime":"2024-01-02T03:04:05Z","mail":null}"""));
        assertThat(user.path("displayName").asText()).isEqualTo("Directory value 1");
        assertThat(user.path("userPrincipalName").asText()).isEqualTo("email-1@example.invalid");
        assertThat(user.path("jobTitle").asText()).isEqualTo("Directory value 2");
        assertThat(user.path("businessPhones").path(0).asText()).isEqualTo("Directory value 3");
        assertThat(user.path("userType").asText()).isEqualTo("Member");
        assertThat(user.path("createdDateTime").asText()).isEqualTo("2024-01-02T03:04:05Z");
        assertThat(user.path("mail").isNull()).isTrue();

        JsonNode group = scrubber.json(Recordings.JSON.readTree("""
                {"Id":5,"LoginName":"Team Owners","Title":"Team Owners","OwnerTitle":"Team Owners",
                 "Description":"Owners of Team","PrincipalType":8}"""));
        assertThat(group.path("Title").asText()).isEqualTo("Principal 1");
        assertThat(group.path("LoginName").asText()).isEqualTo("Principal 1");
        assertThat(group.path("OwnerTitle").asText()).isEqualTo("Principal 1");
        assertThat(group.path("Description").asText()).isEqualTo("Principal 2");
        JsonNode own = scrubber.json(Recordings.JSON.readTree("""
                {"Id":31,"LoginName":"pipestream-live-write-1-group","Title":"pipestream-live-write-1-group",
                 "OwnerTitle":"Team Owners","PrincipalType":8}"""));
        assertThat(own.path("Title").asText()).isEqualTo("pipestream-live-write-1-group");
        assertThat(own.path("LoginName").asText()).isEqualTo("pipestream-live-write-1-group");
        assertThat(own.path("OwnerTitle").asText()).isEqualTo("Principal 1");
        JsonNode made = scrubber.json(Recordings.JSON.readTree("""
                {"id":"g1","displayName":"pipestream-live-write-1-entra",
                 "mailNickname":"pipestream-live-write-1-entra","securityEnabled":true,
                 "description":"pipestream-live-write-1-entra created by a live test run",
                 "securityIdentifier":"S-1-12-1-1"}"""));
        assertThat(made.path("displayName").asText()).isEqualTo("pipestream-live-write-1-entra");
        assertThat(made.path("mailNickname").asText()).isEqualTo("pipestream-live-write-1-entra");
        assertThat(made.path("description").asText())
                .isEqualTo("pipestream-live-write-1-entra created by a live test run");
        assertThat(made.path("securityIdentifier").asText()).isEqualTo("Directory value 4");
        JsonNode member = scrubber.json(Recordings.JSON.readTree("""
                {"Id":9,"LoginName":"i:0#.f|membership|ada@contoso.com","Title":"Ada Lovelace",
                 "UserId":{"NameId":"100320012345abcd","NameIdIssuer":"urn:federation:microsoftonline"}}"""));
        assertThat(member.path("LoginName").asText())
                .isEqualTo("i:0#.f|membership|email-1@example.invalid");
        assertThat(member.path("Title").asText()).isEqualTo("Principal 3");
        assertThat(member.path("UserId").path("NameId").asText()).isEqualTo("nameid-0001");
    }

    @Test
    void aChangedUserOrGroupOfADeltaLosesItsNameWithoutItsOtherProperties() throws Exception {
        JsonNode users = scrubber.json(Recordings.JSON.readTree("""
                {"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#users(id,displayName)",
                 "value":[{"id":"u1","displayName":"Ada Lovelace"},
                          {"id":"u2","@removed":{"reason":"changed"}}]}"""));
        assertThat(users.path("value").path(0).path("displayName").asText())
                .isEqualTo("Directory value 1");
        assertThat(users.path("value").path(1).path("@removed").path("reason").asText())
                .isEqualTo("changed");
        JsonNode groups = scrubber.json(Recordings.JSON.readTree("""
                {"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#groups",
                 "value":[{"id":"g1","description":"Payroll team"}]}"""));
        assertThat(groups.path("value").path(0).path("description").asText())
                .isEqualTo("Directory value 2");
        JsonNode drives = scrubber.json(Recordings.JSON.readTree("""
                {"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#drives",
                 "value":[{"id":"d1","name":"Documents"}]}"""));
        assertThat(drives.path("value").path(0).path("name").asText()).isEqualTo("Documents");
    }

    @Test
    void pagingAndDeltaTokensOfGraphLinksAreReplaced() {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        String link = scrubber.text("https://graph.microsoft.com/v1.0/groups/delta?$deltatoken=AbC-dEf_123.xyz");
        assertThat(link).isEqualTo("https://graph.microsoft.com/v1.0/groups/delta?$deltatoken=cursor-0001");
        assertThat(scrubber.text("https://graph.microsoft.com/v1.0/groups/delta?$skiptoken=Q9z&$top=5"))
                .isEqualTo("https://graph.microsoft.com/v1.0/groups/delta?$skiptoken=cursor-0002&$top=5");
        assertThat(scrubber.text("https://graph.microsoft.com/v1.0/groups/delta?$deltatoken=latest"))
                .endsWith("$deltatoken=latest");
        assertThat(scrubber.text(link)).isEqualTo(link);
    }

    @Test
    void aSitePermissionIdIsReplacedInTheListingAndInLaterUrls() throws Exception {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        String id = RecordingScrubber.ENCODED_GRANTEE + "w5NzFlMDNhMi1hYmNk";
        JsonNode listed = scrubber.json(Recordings.JSON.readTree("""
                {"value":[{"id":"%s","grantedToIdentitiesV2":[{"application":
                  {"id":"0f8fad5b-d9cb-469f-a165-70867728950e","displayName":"Contoso Connector"}}]}]}"""
                .formatted(id)));

        assertThat(listed.path("value").path(0).path("id").asText()).isEqualTo("permission-0001");
        assertThat(scrubber.text("https://graph.microsoft.com/v1.0/sites/s/permissions/" + id))
                .isEqualTo("https://graph.microsoft.com/v1.0/sites/s/permissions/permission-0001");
        assertThat(RecordingScrubber.violations("\"id\" : \"" + id + "\""))
                .containsExactly("encoded grantee of a site permission");
        assertThat(RecordingScrubber.violations(listed.toString())).isEmpty();
    }

    @Test
    void theShownValueOfALookupOrPersonColumnIsReplaced() throws Exception {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        JsonNode fields = scrubber.json(Recordings.JSON.readTree("""
                {"Title":"kept","PsPerson":"Ada Lovelace","PsPersonLookupId":"6",
                 "PsPeople":[{"LookupId":6,"LookupValue":"Ada Lovelace","Email":"ada@contoso.com"}]}"""));

        assertThat(fields.path("Title").asText()).isEqualTo("kept");
        assertThat(fields.path("PsPersonLookupId").asText()).isEqualTo("6");
        assertThat(fields.path("PsPerson").asText()).isEqualTo("Lookup value 1");
        assertThat(fields.path("PsPeople").path(0).path("LookupValue").asText())
                .isEqualTo("Lookup value 1");
        assertThat(fields.toString()).doesNotContain("Ada").doesNotContain("contoso");
    }

    @Test
    void anAddressInsideAUrlIsReplacedAndTheLeakCheckSeesOne() {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");
        String plain = scrubber.text("ada@contoso.com");
        String url = "https://graph.microsoft.com/v1.0/sites/s/lists/l/items?$filter="
                + "fields%2FEMail%20eq%20%27ada%40contoso.com%27";

        String scrubbed = scrubber.text(url);

        assertThat(plain).isEqualTo("email-1@example.invalid");
        assertThat(scrubbed).endsWith("eq%20%27email-1%40example.invalid%27");
        assertThat(RecordingScrubber.violations(url))
                .containsExactly("percent-encoded email address");
        assertThat(RecordingScrubber.violations(scrubbed)).isEmpty();
    }

    @Test
    void aDocumentationUrlIsKept() {
        RecordingScrubber scrubber = new RecordingScrubber("https://graph.microsoft.com/v1.0");

        assertThat(scrubber.text("https://example.com/page, A page"))
                .isEqualTo("https://example.com/page, A page");
    }

    @Test
    void theJoinProofAddressKeepsOnePlaceholderInAnyCase() {
        RecordingScrubber scrubber = new RecordingScrubber(GRAPH);
        scrubber.joinProof("Proof.Person@Contoso.com");
        assertThat(scrubber.text("smtp:proof.person@contoso.com"))
                .isEqualTo("smtp:" + RecordingScrubber.JOIN_PROOF);
        assertThat(scrubber.text("PROOF.PERSON@CONTOSO.COM")).isEqualTo(RecordingScrubber.JOIN_PROOF);
        assertThat(scrubber.text("someone.else@contoso.com")).isEqualTo("email-1@example.invalid");
        assertThat(RecordingScrubber.violations(RecordingScrubber.JOIN_PROOF)).isEmpty();
    }

    @Test
    void aChangeLogEntryKeepsItsKindAndIdsAndLosesNamesTitlesAndPaths() throws Exception {
        JsonNode scrubbed = scrubber.json(Recordings.JSON.readTree("""
                {"value":[{"odata.type":"SP.ChangeItem",
                  "ChangeToken":{"StringValue":"1;3;0f8fad5b-d9cb-469f-a165-70867728950e;638000000000000000;42"},
                  "ChangeType":17,"Time":"2026-11-24T17:00:00Z","ItemId":4,
                  "ListId":"0f8fad5b-d9cb-469f-a165-70867728950e","Editor":"Ada Lovelace",
                  "EditorEmailHint":"ada@contoso.com","ListTitle":"Board papers",
                  "ServerRelativeUrl":"/sites/Board/Lists/Papers/4_.000",
                  "ContentTypeId":{"StringValue":"0x0100AB"},
                  "SharedWithUsers":[{"Name":"Grace Hopper","Email":"grace@contoso.com"}]}]}"""));

        JsonNode change = scrubbed.path("value").get(0);
        String list = change.path("ListId").asText();
        assertThat(change.path("odata.type").asText()).isEqualTo("SP.ChangeItem");
        assertThat(change.path("Time").asText()).isEqualTo("2026-11-24T17:00:00Z");
        assertThat(change.path("ChangeType").asInt()).isEqualTo(17);
        assertThat(change.path("ItemId").asInt()).isEqualTo(4);
        assertThat(list).startsWith(RecordingScrubber.GUID_PREFIX);
        assertThat(change.path("ChangeToken").path("StringValue").asText())
                .isEqualTo("1;3;" + list + ";638000000000000000;42");
        assertThat(change.path("EditorEmailHint").asText()).isEqualTo("email-1@example.invalid");
        assertThat(scrubbed.toString()).doesNotContain("Ada", "Board", "Papers", "Grace", "contoso", "0x0100AB");
        assertThat(RecordingScrubber.violations(scrubbed.toString())).isEmpty();
    }

    @Test
    void xmlSchemaNamespacesStay() {
        String xml = "<Value xmlns:q1=\"http://www.w3.org/2001/XMLSchema\" p4:type=\"q1:string\">";

        assertThat(scrubber.text(xml)).isEqualTo(xml);
    }
}
