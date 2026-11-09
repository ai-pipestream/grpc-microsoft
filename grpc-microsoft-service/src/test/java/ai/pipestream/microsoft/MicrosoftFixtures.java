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

    static String fileJsonWithHashes(String id, String name, String driveId) {
        return """
                {
                  "id": "%s",
                  "name": "%s",
                  "size": 12,
                  "description": "a note",
                  "eTag": "etag-1",
                  "webUrl": "https://contoso.sharepoint.com/%s",
                  "@microsoft.graph.downloadUrl": "https://contoso.sharepoint.com/download/%s",
                  "createdDateTime": "2024-03-01T00:00:00Z",
                  "lastModifiedDateTime": "2024-03-02T00:00:00Z",
                  "file": {
                    "mimeType": "text/plain",
                    "hashes": {
                      "sha1Hash": "aaa",
                      "sha256Hash": "bbb",
                      "quickXorHash": "ccc",
                      "crc32Hash": "ddd"
                    }
                  },
                  "parentReference": {"id": "root", "driveId": "%s"},
                  "createdBy": {"user": {"id": "user-1", "displayName": "Bot"}},
                  "lastModifiedBy": {"user": {"id": "user-1", "displayName": "Bot"}}
                }
                """.formatted(id, name, name, name, driveId);
    }

    static String listItemFieldsJson() {
        return """
                {
                  "fields": {
                    "Title": "Notes",
                    "Count": 3,
                    "Flag": true,
                    "When": "2024-03-02T00:00:00Z",
                    "@odata.etag": "skip-me"
                  }
                }
                """;
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

    /** A delta page that ends the enumeration with {@code deltaLink}. */
    static String deltaJson(String deltaLink, String... itemJsons) {
        return "{\"value\":[" + String.join(",", itemJsons) + "],\"@odata.deltaLink\":\""
                + deltaLink + "\"}";
    }

    /** A delta (or collection) page that continues at {@code nextLink}. */
    static String pageJson(String nextLink, String... itemJsons) {
        return "{\"value\":[" + String.join(",", itemJsons) + "],\"@odata.nextLink\":\""
                + nextLink + "\"}";
    }

    /** A delta delete record (OneDrive for Business shape: no name, no cTag). */
    static String deletedJson(String id, String driveId) {
        return """
                {"id": "%s", "parentReference": {"id": "root", "driveId": "%s"},
                 "file": {}, "deleted": {}}
                """.formatted(id, driveId);
    }

    /** A drive root item as delta returns it. */
    static String rootJson(String driveId) {
        return """
                {"id": "root", "name": "root", "root": {}, "folder": {"childCount": 2},
                 "parentReference": {"driveId": "%s"}}
                """.formatted(driveId);
    }
}
