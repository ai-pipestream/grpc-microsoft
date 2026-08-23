package ai.pipestream.microsoft;

final class MicrosoftFixtures {

    private MicrosoftFixtures() {
    }

    static String meJson() {
        return """
                {
                  "id": "user-1",
                  "displayName": "Bot",
                  "userPrincipalName": "bot@contoso.com",
                  "mail": "bot@contoso.com"
                }
                """;
    }

    static String driveJson(String id, String name) {
        return """
                {
                  "id": "%s",
                  "name": "%s",
                  "driveType": "business",
                  "webUrl": "https://contoso.sharepoint.com/%s"
                }
                """.formatted(id, name, name);
    }

    static String siteJson(String id, String name) {
        return """
                {
                  "id": "%s",
                  "name": "%s",
                  "displayName": "%s",
                  "webUrl": "https://contoso.sharepoint.com/sites/%s"
                }
                """.formatted(id, name, name, name);
    }

    static String fileJson(String id, String name, String driveId) {
        return """
                {
                  "id": "%s",
                  "name": "%s",
                  "size": 12,
                  "webUrl": "https://contoso.sharepoint.com/%s",
                  "createdDateTime": "2024-03-01T00:00:00Z",
                  "lastModifiedDateTime": "2024-03-02T00:00:00Z",
                  "file": {"mimeType": "text/plain"},
                  "parentReference": {"id": "root", "driveId": "%s"},
                  "createdBy": {"user": {"id": "user-1", "displayName": "Bot"}},
                  "lastModifiedBy": {"user": {"id": "user-1", "displayName": "Bot"}}
                }
                """.formatted(id, name, name, driveId);
    }

    static String folderJson(String id, String name, String driveId) {
        return """
                {
                  "id": "%s",
                  "name": "%s",
                  "size": 0,
                  "webUrl": "https://contoso.sharepoint.com/%s",
                  "createdDateTime": "2024-03-01T00:00:00Z",
                  "lastModifiedDateTime": "2024-03-02T00:00:00Z",
                  "folder": {"childCount": 1},
                  "parentReference": {"id": "root", "driveId": "%s"}
                }
                """.formatted(id, name, name, driveId);
    }

    static String childrenJson(String... itemJsons) {
        return "{\"value\":[" + String.join(",", itemJsons) + "]}";
    }

    static String sitesJson(String... siteJsons) {
        return "{\"value\":[" + String.join(",", siteJsons) + "]}";
    }

    static String drivesJson(String... driveJsons) {
        return "{\"value\":[" + String.join(",", driveJsons) + "]}";
    }
}
