package ai.pipestream.microsoft;

import ai.pipestream.microsoft.v1.AuthMode;
import ai.pipestream.microsoft.v1.ChangeOperation;
import ai.pipestream.microsoft.v1.ChangeSource;
import ai.pipestream.microsoft.v1.ColumnDefinition;
import ai.pipestream.microsoft.v1.ColumnType;
import ai.pipestream.microsoft.v1.DeltaDriveItemsRequest;
import ai.pipestream.microsoft.v1.DeltaDriveItemsResponse;
import ai.pipestream.microsoft.v1.DeltaListItemsRequest;
import ai.pipestream.microsoft.v1.DeltaListItemsResponse;
import ai.pipestream.microsoft.v1.DeltaPreferences;
import ai.pipestream.microsoft.v1.DeltaResyncReason;
import ai.pipestream.microsoft.v1.DeltaSitesRequest;
import ai.pipestream.microsoft.v1.DeltaSitesResponse;
import ai.pipestream.microsoft.v1.DownloadContentRequest;
import ai.pipestream.microsoft.v1.DownloadContentResponse;
import ai.pipestream.microsoft.v1.DownloadItemRequest;
import ai.pipestream.microsoft.v1.Drive;
import ai.pipestream.microsoft.v1.DriveItem;
import ai.pipestream.microsoft.v1.DriveItemInclude;
import ai.pipestream.microsoft.v1.DriveItemRef;
import ai.pipestream.microsoft.v1.GetDriveRequest;
import ai.pipestream.microsoft.v1.GetItemRequest;
import ai.pipestream.microsoft.v1.GetListItemRequest;
import ai.pipestream.microsoft.v1.GetListRequest;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetSiteRequest;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.ListAllSitesRequest;
import ai.pipestream.microsoft.v1.ListChildrenRequest;
import ai.pipestream.microsoft.v1.ListColumn;
import ai.pipestream.microsoft.v1.ListColumnsRequest;
import ai.pipestream.microsoft.v1.ListDrivesRequest;
import ai.pipestream.microsoft.v1.ListItem;
import ai.pipestream.microsoft.v1.ListItemInclude;
import ai.pipestream.microsoft.v1.ListItemRef;
import ai.pipestream.microsoft.v1.ListItemVersionsRequest;
import ai.pipestream.microsoft.v1.ListListItemsRequest;
import ai.pipestream.microsoft.v1.ListListsRequest;
import ai.pipestream.microsoft.v1.ListPermissionsRequest;
import ai.pipestream.microsoft.v1.ListSitesRequest;
import ai.pipestream.microsoft.v1.ListSubsitesRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.Permission;
import ai.pipestream.microsoft.v1.PermissionTarget;
import ai.pipestream.microsoft.v1.ProbeConnectionRequest;
import ai.pipestream.microsoft.v1.ProbeConnectionResponse;
import ai.pipestream.microsoft.v1.RetentionBehavior;
import ai.pipestream.microsoft.v1.SharePointList;
import ai.pipestream.microsoft.v1.Site;
import ai.pipestream.microsoft.v1.SyncCursor;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.StatusRuntimeException;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

/**
 * One test per {@code MicrosoftService} rpc against the in-process fake Graph:
 * each asserts the exact Graph request (path, query, {@code Prefer}) and the
 * typed response.
 */
class MicrosoftRpcSurfaceTest {

    private static final String SITE = "contoso.sharepoint.com,c1,w1";

    private FakeGraphServer fake;
    private final List<Server> servers = new ArrayList<>();
    private final List<ManagedChannel> channels = new ArrayList<>();
    private final ExecutorService executor = Executors.newCachedThreadPool();
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub;
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub appStub;
    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub delegatedStub;

    @BeforeEach
    void start() throws Exception {
        fake = FakeGraphServer.start();
        stub = startService(AuthMode.AUTH_MODE_UNSPECIFIED);
        appStub = startService(AuthMode.AUTH_MODE_APPLICATION);
        delegatedStub = startService(AuthMode.AUTH_MODE_DELEGATED);
    }

    private MicrosoftServiceGrpc.MicrosoftServiceBlockingStub startService(AuthMode mode)
            throws Exception {
        MicrosoftConnectorConfig config = MicrosoftConnectorConfig.builder()
                .tenantId("tenant-1").clientId("client-1").clientSecret("env-secret")
                .graphBaseUrl(fake.baseUrl()).authority(fake.baseUrl())
                .build();
        GraphFiles files = new GraphFiles(new GraphClient(fake.baseUrl(), () -> "token", mode));
        String name = InProcessServerBuilder.generateName();
        servers.add(InProcessServerBuilder.forName(name).executor(executor)
                .addService(new MicrosoftGrpcService(config, files, 0, null, null, 4))
                .build().start());
        ManagedChannel channel = InProcessChannelBuilder.forName(name).build();
        channels.add(channel);
        return MicrosoftServiceGrpc.newBlockingStub(channel);
    }

    @AfterEach
    void stop() {
        channels.forEach(ManagedChannel::shutdownNow);
        servers.forEach(Server::shutdownNow);
        executor.shutdownNow();
        fake.close();
    }

    private static <T> List<T> drain(Iterator<T> it) {
        List<T> out = new ArrayList<>();
        it.forEachRemaining(out::add);
        return out;
    }

    private static Status.Code code(Runnable call) {
        try {
            call.run();
        } catch (StatusRuntimeException e) {
            return e.getStatus().getCode();
        }
        return Status.Code.OK;
    }

    // ------------------------------------------------------------ identity

    @Test
    void getMeSelectsEveryModelledUserPropertyAndRefusesAppTokens() {
        fake.stub("/me", """
                {"id":"u1","displayName":"Ada","userPrincipalName":"ada@contoso.com",
                 "mail":"ada@contoso.com","givenName":"Ada","surname":"Lovelace","jobTitle":"Eng",
                 "department":"R&D","officeLocation":"B1","userType":"Guest","accountEnabled":false,
                 "preferredLanguage":"en-GB"}
                """);
        GraphUser user = delegatedStub.getMe(GetMeRequest.getDefaultInstance()).getUser();
        assertThat(fake.only("/me").query()).isEqualTo("$select=" + GraphPaths.USER_SELECT);
        assertThat(user.getSurname()).isEqualTo("Lovelace");
        assertThat(user.getUserType()).isEqualTo("Guest");
        assertThat(user.hasAccountEnabled()).isTrue();
        assertThat(user.getAccountEnabled()).isFalse();
        assertThat(user.getPreferredLanguage()).isEqualTo("en-GB");

        assertThat(code(() -> appStub.getMe(GetMeRequest.getDefaultInstance())))
                .isEqualTo(Status.Code.FAILED_PRECONDITION);
        assertThat(fake.requestsTo("/me")).hasSize(1);
    }

    @Test
    void probeConnectionReportsAuthModeAndNeverEchoesAnInlineSecret() {
        fake.stub("/sites/getAllSites", MicrosoftFixtures.sitesJson(
                MicrosoftFixtures.siteJson("s1", "A"), MicrosoftFixtures.siteJson("s2", "B"),
                MicrosoftFixtures.siteJson("s3", "C")));
        ProbeConnectionResponse ok = appStub.probeConnection(ProbeConnectionRequest.newBuilder()
                .setLimit(2).build());
        assertThat(fake.only("/sites/getAllSites").query()).isEqualTo("$select=id");
        assertThat(ok.getOk()).isTrue();
        assertThat(ok.getAuthMode()).isEqualTo(AuthMode.AUTH_MODE_APPLICATION);
        assertThat(ok.getSiteIdsList()).containsExactly("s1", "s2");
        assertThat(ok.getConnectionId()).isEqualTo("default");

        fake.stub("/tenant-x/oauth2/v2.0/token", new FakeGraphServer.Stub(401,
                "{\"error\":\"invalid_client\",\"error_description\":\"AADSTS7000215 bad secret\"}"
                        .getBytes(StandardCharsets.UTF_8), Map.of("content-type", "application/json")));
        ProbeConnectionResponse failed = stub.probeConnection(ProbeConnectionRequest.newBuilder()
                .setTenantId("tenant-x").setClientId("app-x").setClientSecret("s3cr3t-value").build());
        assertThat(failed.getOk()).isFalse();
        assertThat(failed.getConnectionId()).isEmpty();
        assertThat(failed.getAuthMode()).isEqualTo(AuthMode.AUTH_MODE_APPLICATION);
        assertThat(failed.getErrorMessage()).contains("invalid_client").doesNotContain("s3cr3t-value");
        assertThat(failed.toString()).doesNotContain("s3cr3t-value");

        assertThat(code(() -> stub.probeConnection(ProbeConnectionRequest.newBuilder()
                .setTenantId("tenant-x").build()))).isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    // ------------------------------------------------------------ sites

    @Test
    void listSitesWalksEveryPageWhenLimitIsZero() {
        fake.stub("/sites", MicrosoftFixtures.pageJson(fake.baseUrl() + "/sites?search=fin&$skiptoken=2",
                MicrosoftFixtures.siteJson("s1", "A")));
        fake.stub("/sites?search=fin&$skiptoken=2", MicrosoftFixtures.sitesJson(
                MicrosoftFixtures.siteJson("s2", "B")));
        List<Site> sites = stub.listSites(ListSitesRequest.newBuilder().setQuery("fin").build())
                .getSitesList();
        assertThat(fake.requestsTo("/sites").get(0).query())
                .isEqualTo("search=fin&$select=" + GraphPaths.SITE_SELECT);
        assertThat(sites).extracting(Site::getId).containsExactly("s1", "s2");
    }

    @Test
    void listAllSitesStreamsAndFiltersPersonalSites() {
        fake.stub("/sites/getAllSites", """
                {"value":[
                  {"id":"s1","name":"Root","isPersonalSite":false,"root":{},
                   "siteCollection":{"hostName":"contoso.sharepoint.com","dataLocationCode":"NAM","root":{}}},
                  {"id":"p1","name":"Ada","isPersonalSite":true}]}
                """);
        List<Site> sites = drain(appStub.listAllSites(ListAllSitesRequest.getDefaultInstance()))
                .stream().map(r -> r.getSite()).toList();
        assertThat(fake.only("/sites/getAllSites").query()).isEqualTo("$select=" + GraphPaths.SITE_SELECT);
        assertThat(sites).extracting(Site::getId).containsExactly("s1");
        assertThat(sites.get(0).getIsRoot()).isTrue();
        assertThat(sites.get(0).getSiteCollection().getHostname()).isEqualTo("contoso.sharepoint.com");
        assertThat(sites.get(0).getSiteCollection().getIsRoot()).isTrue();
        assertThat(drain(appStub.listAllSites(ListAllSitesRequest.newBuilder()
                .setIncludePersonalSites(true).build()))).hasSize(2);
        assertThat(code(() -> drain(delegatedStub.listAllSites(ListAllSitesRequest.getDefaultInstance()))))
                .isEqualTo(Status.Code.FAILED_PRECONDITION);
    }

    @Test
    void getSiteAcceptsTheHostnamePathForm() {
        fake.stub("/sites/contoso.sharepoint.com:/sites/finance", """
                {"id":"%s","name":"finance","displayName":"Finance","isPersonalSite":false,
                 "eTag":"e1","createdDateTime":"2024-01-01T00:00:00Z",
                 "sharepointIds":{"siteId":"c1","webId":"w1","tenantId":"t1",
                   "siteUrl":"https://contoso.sharepoint.com/sites/finance"},
                 "siteCollection":{"hostname":"contoso.sharepoint.com",
                   "archivalDetails":{"archiveStatus":"recentlyArchived"}}}
                """.formatted(SITE));
        Site site = stub.getSite(GetSiteRequest.newBuilder()
                .setSiteId("contoso.sharepoint.com:/sites/finance").build()).getSite();
        assertThat(fake.only("/sites/contoso.sharepoint.com:/sites/finance").query())
                .isEqualTo("$select=" + GraphPaths.SITE_SELECT);
        assertThat(site.getId()).isEqualTo(SITE);
        assertThat(site.getSharepointIds().getWebId()).isEqualTo("w1");
        assertThat(site.getSiteCollection().getArchivalState()).isEqualTo("recentlyArchived");
        assertThat(site.getEtag()).isEqualTo("e1");
        assertThat(site.hasCreatedAt()).isTrue();
    }

    @Test
    void listSubsitesFillsTheParentSite() {
        fake.stub("/sites/" + SITE + "/sites", MicrosoftFixtures.sitesJson(
                MicrosoftFixtures.siteJson("sub-1", "Team")));
        List<Site> subs = drain(stub.listSubsites(ListSubsitesRequest.newBuilder()
                .setSiteId(SITE).build())).stream().map(r -> r.getSite()).toList();
        assertThat(fake.only("/sites/" + SITE + "/sites").query())
                .isEqualTo("$select=" + GraphPaths.SITE_SELECT);
        assertThat(subs).extracting(Site::getId).containsExactly("sub-1");
        assertThat(subs.get(0).getParentReference().getSiteId()).isEqualTo(SITE);
    }

    @Test
    void deltaSitesEnumeratesResumesResyncsAndValidatesLinks() {
        String d1 = fake.baseUrl() + "/sites/delta?$deltatoken=abc";
        fake.stub("/sites/delta", MicrosoftFixtures.deltaJson(d1, """
                {"id":"s1","name":"A"}""", """
                {"id":"s2","name":"B","deleted":{"state":"deleted"}}"""));
        List<DeltaSitesResponse> full = drain(appStub.deltaSites(DeltaSitesRequest.getDefaultInstance()));
        assertThat(fake.requestsTo("/sites/delta").get(0).query())
                .isEqualTo("$select=" + GraphPaths.SITE_SELECT);
        assertThat(full).hasSize(3);
        assertThat(full.get(1).getSite().getDeleted().getState()).isEqualTo("deleted");
        assertThat(full.get(2).getDeltaLink()).isEqualTo(d1);

        fake.stub("/sites/delta?$deltatoken=abc", FakeGraphServer.error(410,
                "resyncChangesUploadDifferences", Map.of("Location", fake.baseUrl() + "/sites/delta?token=new")));
        fake.stub("/sites/delta?token=new", MicrosoftFixtures.deltaJson(d1, "{\"id\":\"s3\"}"));
        List<DeltaSitesResponse> resumed = drain(appStub.deltaSites(DeltaSitesRequest.newBuilder()
                .setDeltaLink(d1).build()));
        assertThat(resumed.get(0).getResync().getReason())
                .isEqualTo(DeltaResyncReason.DELTA_RESYNC_REASON_UPLOAD_DIFFERENCES);
        assertThat(resumed.get(1).getSite().getId()).isEqualTo("s3");

        fake.stub("/sites/delta?token=latest", MicrosoftFixtures.deltaJson(d1));
        List<DeltaSitesResponse> latest = drain(appStub.deltaSites(DeltaSitesRequest.newBuilder()
                .setDeltaLink("latest").build()));
        assertThat(latest).hasSize(1);
        assertThat(latest.get(0).getDeltaLink()).isEqualTo(d1);

        int before = fake.requests().size();
        assertThat(code(() -> drain(appStub.deltaSites(DeltaSitesRequest.newBuilder()
                .setDeltaLink("https://attacker.example/v1.0/sites/delta?token=x").build()))))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(fake.requests()).hasSize(before);
    }

    // ------------------------------------------------------------ drives

    @Test
    void listDrivesCoversSiteUserGroupAndMeScopes() {
        fake.stub("/groups/g1/drives", """
                {"value":[{"id":"d1","name":"Docs","driveType":"documentLibrary",
                  "owner":{"group":{"id":"g1","displayName":"Team"}},
                  "quota":{"total":100,"used":40,"remaining":60,"deleted":2,"state":"normal"},
                  "sharePointIds":{"listId":"L1","siteId":"c1"},"system":{}}]}
                """);
        fake.stub("/users/ada@contoso.com/drives", MicrosoftFixtures.drivesJson(
                MicrosoftFixtures.driveJson("d2", "OneDrive")));
        Drive group = stub.listDrives(ListDrivesRequest.newBuilder().setGroupId("g1").build())
                .getDrives(0);
        assertThat(fake.only("/groups/g1/drives").query()).isEqualTo("$select=" + GraphPaths.DRIVE_SELECT);
        assertThat(group.getOwner().getGroup().getDisplayName()).isEqualTo("Team");
        assertThat(group.getQuota().getDeleted()).isEqualTo(2);
        assertThat(group.getListId()).isEqualTo("L1");
        assertThat(group.getIsSystem()).isTrue();
        assertThat(stub.listDrives(ListDrivesRequest.newBuilder().setUserId("ada@contoso.com").build())
                .getDrivesList()).extracting(Drive::getId).containsExactly("d2");
        assertThat(code(() -> appStub.listDrives(ListDrivesRequest.getDefaultInstance())))
                .isEqualTo(Status.Code.FAILED_PRECONDITION);
    }

    @Test
    void getDriveSelectsSharePointIds() {
        fake.stub("/drives/d1", MicrosoftFixtures.driveJson("d1", "Docs"));
        Drive drive = stub.getDrive(GetDriveRequest.newBuilder().setDriveId("d1").build()).getDrive();
        assertThat(fake.only("/drives/d1").query()).isEqualTo("$select=" + GraphPaths.DRIVE_SELECT);
        assertThat(drive.getDriveType()).isEqualTo("business");
    }

    @Test
    void listChildrenByFolderIdWithPermissions() {
        fake.stub("/drives/d1/items/f1/children", MicrosoftFixtures.childrenJson(
                MicrosoftFixtures.fileJson("x1", "a.txt", "d1")));
        fake.stub("/drives/d1/items/x1/permissions", """
                {"value":[{"id":"p1","roles":["read"],
                  "grantedToV2":{"user":{"id":"u1","displayName":"Ada","email":"ada@contoso.com"}},
                  "inheritedFrom":{"id":"f1","driveId":"d1"},
                  "expirationDateTime":"0001-01-01T00:00:00Z"}]}
                """);
        List<DriveItem> items = drain(stub.listChildren(ListChildrenRequest.newBuilder()
                .setDriveId("d1").setFolderId("f1")
                .setInclude(DriveItemInclude.newBuilder().setPermissions(true)).build()))
                .stream().map(r -> r.getItem()).toList();
        assertThat(fake.only("/drives/d1/items/f1/children").query())
                .isEqualTo("$select=" + GraphPaths.DRIVE_ITEM_SELECT);
        assertThat(fake.requestsTo("/drives/d1/items/x1/listItem")).isEmpty();
        Permission permission = items.get(0).getPermissions(0);
        assertThat(permission.getGrantedTo().getUser().getEmail()).isEqualTo("ada@contoso.com");
        assertThat(permission.getInheritedFrom().getId()).isEqualTo("f1");
        assertThat(permission.hasExpiresAt()).isFalse();
        assertThat(permission.getTarget().getDriveItem().getItemId()).isEqualTo("x1");
    }

    @Test
    void getItemByPathWithRetentionLabelAndPublication() {
        fake.stub("/drives/d1/root:/Shared Documents/report.docx:", """
                {"id":"x1","name":"report.docx","size":5,"cTag":"c1","webDavUrl":"https://dav",
                 "file":{"mimeType":"application/msword","hashes":{"quickXorHash":"q"}},
                 "publication":{"level":"checkout","versionId":"2.0",
                   "checkedOutBy":{"user":{"id":"u1"}}},
                 "sharepointIds":{"listId":"L1","listItemId":"9"},
                 "shared":{"scope":"organization"},"malware":{"description":"eicar"},
                 "pendingOperations":{"pendingContentUpdate":{"queuedDateTime":"2024-01-01T00:00:00Z"}},
                 "fileSystemInfo":{"createdDateTime":"2023-01-01T00:00:00Z"},
                 "image":{"width":10,"height":20},
                 "photo":{"cameraMake":"X","iso":100,"takenDateTime":"2023-01-01T00:00:00Z"},
                 "location":{"latitude":1.5,"longitude":2.5},
                 "parentReference":{"id":"root","driveId":"d1","siteId":"s1","driveType":"documentLibrary"}}
                """);
        fake.stub("/drives/d1/items/x1/retentionLabel", """
                {"name":"Keep 7y","isLabelAppliedExplicitly":true,
                 "labelAppliedDateTime":"2024-01-01T00:00:00Z",
                 "retentionSettings":{"behaviorDuringRetentionPeriod":"retainAsRecord",
                   "isDeleteAllowed":false,"isRecordLocked":true}}
                """);
        DriveItem item = stub.getItem(GetItemRequest.newBuilder().setDriveId("d1")
                .setPath("/Shared Documents/report.docx")
                .setInclude(DriveItemInclude.newBuilder().setRetentionLabel(true).setPublication(true))
                .build()).getItem();
        assertThat(fake.only("/drives/d1/root:/Shared Documents/report.docx:").query())
                .isEqualTo("$select=" + GraphPaths.DRIVE_ITEM_SELECT + ",publication");
        assertThat(item.getPublication().getLevel()).isEqualTo("checkout");
        assertThat(item.getPublication().getCheckedOutBy().getUser().getId()).isEqualTo("u1");
        assertThat(item.getRetentionLabel().getRetentionSettings().getBehaviorDuringRetentionPeriod())
                .isEqualTo(RetentionBehavior.RETENTION_BEHAVIOR_RETAIN_AS_RECORD);
        assertThat(item.getRetentionLabel().getRetentionSettings().getIsRecordLocked()).isTrue();
        assertThat(item.getCtag()).isEqualTo("c1");
        assertThat(item.getListItemId()).isEqualTo("9");
        assertThat(item.getShared().getScope()).isEqualTo("organization");
        assertThat(item.getMalware().getDescription()).isEqualTo("eicar");
        assertThat(item.getPendingContentUpdate()).isTrue();
        assertThat(item.getPhoto().getIso()).isEqualTo(100);
        assertThat(item.getLocation().getLatitude()).isEqualTo(1.5);
        assertThat(item.getParentReference().getSiteId()).isEqualTo("s1");
        assertThat(item.getHashes().getQuickXor()).isEqualTo("q");
        assertThat(fake.requestsTo("/drives/d1/items/x1/listItem")).isEmpty();
    }

    @Test
    @SuppressWarnings("deprecation")
    void downloadItemHonoursTheRequestInlineLimit() {
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));
        fake.stub("/drives/d1/items/x1/content", FakeGraphServer.Stub.bytes(
                "hello world!".getBytes(StandardCharsets.UTF_8), "text/plain"));
        assertThat(code(() -> stub.downloadItem(DownloadItemRequest.newBuilder().setDriveId("d1")
                .setItemId("x1").setIncludeContent(true).setMaxInlineBytes(5).build())))
                .isEqualTo(Status.Code.FAILED_PRECONDITION);
        DriveItem item = stub.downloadItem(DownloadItemRequest.newBuilder().setDriveId("d1")
                .setItemId("x1").setIncludeContent(true).build()).getItem();
        assertThat(item.getContent().toStringUtf8()).isEqualTo("hello world!");
        assertThat(fake.requestsTo("/drives/d1/items/x1/listItem")).isNotEmpty();
    }

    @Test
    void downloadContentStreamsMetadataThenChunksAndKeepsTheTokenOffTheRedirect() throws Exception {
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));
        fake.stub("/drives/d1/items/x1/versions/2.0/content", new FakeGraphServer.Stub(302, new byte[0],
                Map.of("Location", fake.baseUrl() + "/preauth/blob")));
        fake.stub("/preauth/blob", new FakeGraphServer.Stub(206,
                "cdefghij".getBytes(StandardCharsets.UTF_8), Map.of("content-type", "text/plain")));
        List<DownloadContentResponse> stream = drain(stub.downloadContent(DownloadContentRequest.newBuilder()
                .setDriveId("d1").setItemId("x1").setVersionId("2.0").setFormat("pdf")
                .setOffset(2).setLength(8).setChunkBytes(3).build()));
        assertThat(stream.get(0).getItem().getId()).isEqualTo("x1");
        assertThat(stream.subList(1, stream.size())).extracting(r -> r.getChunk().toStringUtf8())
                .containsExactly("cde", "fgh", "ij");
        FakeGraphServer.RecordedRequest content = fake.only("/drives/d1/items/x1/versions/2.0/content");
        assertThat(content.query()).isEqualTo("format=pdf");
        assertThat(content.authorization()).isEqualTo("Bearer token");
        assertThat(content.range()).isNull();
        FakeGraphServer.RecordedRequest blob = fake.only("/preauth/blob");
        assertThat(blob.authorization()).isNull();
        assertThat(blob.range()).isEqualTo("bytes=2-9");

        fake.stub("/drives/d1/items/x1/content", FakeGraphServer.Stub.bytes(
                "0123456789".getBytes(StandardCharsets.UTF_8), "text/plain"));
        ByteArrayOutputStream bytes = new ByteArrayOutputStream();
        List<DownloadContentResponse> sliced = drain(stub.downloadContent(DownloadContentRequest
                .newBuilder().setDriveId("d1").setItemId("x1").setOffset(3).setLength(4).build()));
        for (DownloadContentResponse r : sliced.subList(1, sliced.size())) {
            bytes.write(r.getChunk().toByteArray());
        }
        assertThat(bytes.toString(StandardCharsets.UTF_8)).isEqualTo("3456");
        assertThat(sliced).hasSize(2);
    }

    @Test
    void downloadContentOffsetAtOrPastTheEndIsOutOfRange() {
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));
        fake.stub("/drives/d1/items/x1/content", FakeGraphServer.Stub.bytes(
                "0123456789".getBytes(StandardCharsets.UTF_8), "text/plain"));
        for (long offset : new long[] {10, 12}) {
            assertThat(code(() -> drain(stub.downloadContent(DownloadContentRequest.newBuilder()
                    .setDriveId("d1").setItemId("x1").setOffset(offset).build()))))
                    .as("offset %d", offset).isEqualTo(Status.Code.OUT_OF_RANGE);
        }
        List<DownloadContentResponse> last = drain(stub.downloadContent(DownloadContentRequest
                .newBuilder().setDriveId("d1").setItemId("x1").setOffset(9).build()));
        assertThat(last.get(1).getChunk().toStringUtf8()).isEqualTo("9");
    }

    @Test
    void deltaDriveItemsKeepsGoingWhenAnItemIsGoneBeforeItsPermissionsAreRead() {
        String link = fake.baseUrl() + "/drives/d1/root/delta?token=n";
        fake.stub("/drives/d1/root/delta", MicrosoftFixtures.deltaJson(link,
                MicrosoftFixtures.fileJson("x1", "a.txt", "d1")));
        // /drives/d1/items/x1/permissions is not stubbed: the fake answers 404.
        List<DeltaDriveItemsResponse> events = drain(stub.deltaDriveItems(DeltaDriveItemsRequest
                .newBuilder().setDriveId("d1")
                .setInclude(DriveItemInclude.newBuilder().setPermissions(true)).build()));
        assertThat(fake.requestsTo("/drives/d1/items/x1/permissions")).hasSize(1);
        assertThat(events.get(0).getItem().getPermissionsList()).isEmpty();
        assertThat(events.getLast().getDeltaLink()).isEqualTo(link);
    }

    @Test
    void streamsWaitForTheClientAndStopPagingWhenItCancels() throws Exception {
        for (int i = 0; i < 5; i++) {
            String item = MicrosoftFixtures.fileJson("x" + i, "f" + i + ".txt", "d1");
            String body = i == 4 ? MicrosoftFixtures.deltaJson(fake.baseUrl() + "/drives/d1/root/delta?token=end", item)
                    : MicrosoftFixtures.pageJson(fake.baseUrl() + "/drives/d1/root/delta?page=" + (i + 1), item);
            fake.stub(i == 0 ? "/drives/d1/root/delta" : "/drives/d1/root/delta?page=" + i, body);
        }
        java.util.concurrent.CountDownLatch first = new java.util.concurrent.CountDownLatch(1);
        java.util.concurrent.CountDownLatch closed = new java.util.concurrent.CountDownLatch(1);
        java.util.concurrent.atomic.AtomicReference<io.grpc.stub.ClientCallStreamObserver<DeltaDriveItemsRequest>>
                call = new java.util.concurrent.atomic.AtomicReference<>();
        MicrosoftServiceGrpc.newStub(channels.get(0)).deltaDriveItems(
                DeltaDriveItemsRequest.newBuilder().setDriveId("d1").build(),
                new io.grpc.stub.ClientResponseObserver<DeltaDriveItemsRequest, DeltaDriveItemsResponse>() {
                    @Override
                    public void beforeStart(
                            io.grpc.stub.ClientCallStreamObserver<DeltaDriveItemsRequest> requestStream) {
                        call.set(requestStream);
                        requestStream.disableAutoRequestWithInitial(1);
                    }

                    @Override
                    public void onNext(DeltaDriveItemsResponse value) {
                        first.countDown();
                    }

                    @Override
                    public void onError(Throwable t) {
                        closed.countDown();
                    }

                    @Override
                    public void onCompleted() {
                        closed.countDown();
                    }
                });
        assertThat(first.await(5, java.util.concurrent.TimeUnit.SECONDS)).isTrue();
        Thread.sleep(300);
        assertThat(fake.requestsTo("/drives/d1/root/delta")).hasSizeLessThanOrEqualTo(2);
        call.get().cancel("done", null);
        assertThat(closed.await(5, java.util.concurrent.TimeUnit.SECONDS)).isTrue();
        Thread.sleep(300);
        assertThat(fake.requestsTo("/drives/d1/root/delta")).hasSizeLessThanOrEqualTo(2);
    }

    @Test
    void aRepeatedNextLinkIsInternalNotASpin() {
        String again = fake.baseUrl() + "/sites/" + SITE + "/drives?page=again";
        fake.stub("/sites/" + SITE + "/drives", MicrosoftFixtures.pageJson(again,
                "{\"id\":\"d1\",\"driveType\":\"documentLibrary\"}"));
        fake.stub("/sites/" + SITE + "/drives?page=again", MicrosoftFixtures.pageJson(again,
                "{\"id\":\"d2\",\"driveType\":\"documentLibrary\"}"));
        assertThat(code(() -> stub.listDrives(ListDrivesRequest.newBuilder().setSiteId(SITE).build())))
                .isEqualTo(Status.Code.INTERNAL);
        assertThat(fake.requestsTo("/sites/" + SITE + "/drives")).hasSize(2);
    }

    @Test
    void deltaDriveItemsSendsPreferHeadersAndFlagsSharingChanges() {
        String d1 = fake.baseUrl() + "/drives/d1/root/delta(token='abc')";
        fake.stub("/drives/d1/root/delta", MicrosoftFixtures.deltaJson(d1, """
                {"id":"x1","name":"a.txt","file":{},"@microsoft.graph.sharedChanged":"True",
                 "shared":{},"parentReference":{"driveId":"d1"}}""",
                MicrosoftFixtures.deletedJson("x2", "d1")));
        List<DeltaDriveItemsResponse> events = drain(stub.deltaDriveItems(DeltaDriveItemsRequest
                .newBuilder().setDriveId("d1")
                .setPreferences(DeltaPreferences.newBuilder().setShowRemovedAsDeleted(true)
                        .setTraversePermissionGaps(true).setShowSharingChanges(true)
                        .setExcludeParent(true))
                .build()));
        FakeGraphServer.RecordedRequest request = fake.only("/drives/d1/root/delta");
        assertThat(request.query()).isEqualTo("$select=" + GraphPaths.DRIVE_ITEM_SELECT);
        assertThat(request.prefer()).isEqualTo("deltashowremovedasdeleted, deltatraversepermissiongaps, "
                + "deltashowsharingchanges, deltaExcludeParent");
        assertThat(events.get(0).getItem().getSharingChanged()).isTrue();
        assertThat(events.get(0).getItem().hasShared()).isTrue();
        assertThat(events.get(1).getItem().hasDeleted()).isTrue();
        assertThat(events.get(2).getDeltaLink()).isEqualTo(d1);
        assertThat(fake.requestsTo("/drives/d1/items/x1/listItem")).isEmpty();

        assertThat(code(() -> drain(stub.deltaDriveItems(DeltaDriveItemsRequest.newBuilder()
                .setDriveId("d1").setPreferences(DeltaPreferences.newBuilder().setShowSharingChanges(true))
                .build())))).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(code(() -> drain(stub.deltaDriveItems(DeltaDriveItemsRequest.newBuilder()
                .setDriveId("d1").setDeltaLink(fake.baseUrl() + "/drives/other/root/delta?token=x")
                .build())))).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(fake.requestsTo("/drives/other/root/delta")).isEmpty();
    }

    @Test
    void deltaDriveItemsResyncsOnGoneAndRefusesFolderScopeOnBusinessDrives() {
        String stale = fake.baseUrl() + "/drives/d1/root/delta?token=stale";
        fake.stub("/drives/d1/root/delta?token=stale", FakeGraphServer.error(410,
                "resyncChangesApplyDifferences", Map.of("Location", fake.baseUrl() + "/drives/d1/root/delta?token=r")));
        fake.stub("/drives/d1/root/delta?token=r", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/d1/root/delta?token=n",
                MicrosoftFixtures.fileJson("x1", "a.txt", "d1")));
        List<DeltaDriveItemsResponse> events = drain(stub.deltaDriveItems(DeltaDriveItemsRequest
                .newBuilder().setDriveId("d1").setDeltaLink(stale).build()));
        assertThat(events.get(0).getResync().getReason())
                .isEqualTo(DeltaResyncReason.DELTA_RESYNC_REASON_APPLY_DIFFERENCES);
        assertThat(events.get(1).getItem().getId()).isEqualTo("x1");
        assertThat(events.get(2).getDeltaLink()).endsWith("token=n");

        fake.stub("/drives/d1", "{\"id\":\"d1\",\"driveType\":\"business\"}");
        assertThat(code(() -> drain(stub.deltaDriveItems(DeltaDriveItemsRequest.newBuilder()
                .setDriveId("d1").setItemId("folder-9").build())))).isEqualTo(Status.Code.INVALID_ARGUMENT);
        fake.stub("/drives/p1", "{\"id\":\"p1\",\"driveType\":\"personal\"}");
        fake.stub("/drives/p1/items/folder-9/delta", MicrosoftFixtures.deltaJson(
                fake.baseUrl() + "/drives/p1/items/folder-9/delta?token=1"));
        assertThat(drain(stub.deltaDriveItems(DeltaDriveItemsRequest.newBuilder()
                .setDriveId("p1").setItemId("folder-9").build()))).hasSize(1);
    }

    @Test
    void listItemVersionsSetsBackReferences() {
        fake.stub("/drives/d1/items/x1/versions", """
                {"value":[{"id":"2.0","size":10,"lastModifiedDateTime":"2024-01-02T00:00:00Z",
                  "lastModifiedBy":{"user":{"id":"u1"}},"publication":{"level":"published","versionId":"2.0"}}]}
                """);
        var version = drain(stub.listItemVersions(ListItemVersionsRequest.newBuilder()
                .setDriveId("d1").setItemId("x1").build())).get(0).getVersion();
        assertThat(version.getId()).isEqualTo("2.0");
        assertThat(version.getDriveId()).isEqualTo("d1");
        assertThat(version.getItemId()).isEqualTo("x1");
        assertThat(version.getPublication().getLevel()).isEqualTo("published");
        assertThat(version.getSize()).isEqualTo(10);
    }

    // ------------------------------------------------------------ lists

    @Test
    void listListsFiltersHiddenAndSystemLists() {
        fake.stub("/sites/" + SITE + "/lists", """
                {"value":[
                  {"id":"L1","name":"Tasks","displayName":"Tasks","list":{"template":"genericList"}},
                  {"id":"L2","name":"Hidden","list":{"hidden":true,"template":"genericList"}},
                  {"id":"L3","name":"Sys","system":{},"list":{"template":"genericList"}}]}
                """);
        List<SharePointList> lists = drain(stub.listLists(ListListsRequest.newBuilder()
                .setSiteId(SITE).build())).stream().map(r -> r.getList()).toList();
        assertThat(fake.requestsTo("/sites/" + SITE + "/lists").get(0).query())
                .isEqualTo("$select=" + GraphPaths.LIST_SELECT);
        assertThat(lists).extracting(SharePointList::getId).containsExactly("L1");
        assertThat(lists.get(0).getSiteId()).isEqualTo(SITE);
        assertThat(drain(stub.listLists(ListListsRequest.newBuilder().setSiteId(SITE)
                .setIncludeHidden(true).build()))).hasSize(3);
    }

    @Test
    void getListBehindADocumentLibrary() {
        fake.stub("/drives/d1/list", """
                {"id":"L1","name":"Shared Documents","list":{"template":"documentLibrary",
                 "contentTypesEnabled":true},"parentReference":{"siteId":"%s"}}
                """.formatted(SITE));
        SharePointList list = stub.getList(GetListRequest.newBuilder().setDriveId("d1").build()).getList();
        assertThat(fake.only("/drives/d1/list").query()).isEqualTo("$select=" + GraphPaths.LIST_SELECT);
        assertThat(list.getDriveId()).isEqualTo("d1");
        assertThat(list.getSiteId()).isEqualTo(SITE);
        assertThat(list.getList().getContentTypesEnabled()).isTrue();
        assertThat(code(() -> stub.getList(GetListRequest.getDefaultInstance())))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void listListItemsFiltersWithTheNonIndexedPreferenceAndTypesEveryColumn() {
        String route = "/sites/" + SITE + "/lists/L1/items";
        fake.stub(route, """
                {"value":[{"id":"1","contentType":{"id":"0x01","name":"Item"},
                  "fields":{"Title":"Plan","Status":"Active","Due":"2024-05-01T00:00:00Z",
                    "Tags":["a","b"],
                    "Owners":[{"LookupId":4,"LookupValue":"Ada","Email":"ada@contoso.com"},
                              {"LookupId":5,"LookupValue":"Bob"}],
                    "Place":{"DisplayName":"HQ","Address":{"City":"Oslo"}}}}]}
                """);
        List<ListItem> items = drain(stub.listListItems(ListListItemsRequest.newBuilder()
                .setSiteId(SITE).setListId("L1").setFilter("fields/Status eq 'Active'").build()))
                .stream().map(r -> r.getItem()).toList();
        FakeGraphServer.RecordedRequest request = fake.only(route);
        assertThat(request.query()).isEqualTo("$select=" + GraphPaths.LIST_ITEM_SELECT
                + "&$expand=fields&$filter=fields%2FStatus%20eq%20%27Active%27");
        assertThat(request.prefer()).isEqualTo("HonorNonIndexedQueriesWarningMayFailRandomly");
        ListItem item = items.get(0);
        assertThat(item.getSiteId()).isEqualTo(SITE);
        assertThat(item.getListId()).isEqualTo("L1");
        assertThat(item.getContentType().getName()).isEqualTo("Item");
        assertThat(column(item, "Due").hasTimestampValue()).isTrue();
        assertThat(column(item, "Tags").getStringValuesList()).containsExactly("a", "b");
        assertThat(column(item, "Owners.LookupValue").getStringValuesList()).containsExactly("Ada", "Bob");
        assertThat(column(item, "Owners.LookupId").getStringValuesList()).containsExactly("4", "5");
        assertThat(column(item, "Owners.Email").getStringValuesList()).containsExactly("ada@contoso.com");
        assertThat(column(item, "Place.Address.City").getStringValue()).isEqualTo("Oslo");
    }

    @Test
    void getListItemExpandsTheDriveItemAndPermissions() {
        String route = "/sites/" + SITE + "/lists/L1/items/3";
        fake.stub(route, """
                {"id":"3","fields":{"Title":"Doc"},
                 "driveItem":{"id":"x1","name":"doc.docx","parentReference":{"driveId":"d1"}}}
                """);
        fake.stub(route + "/permissions", """
                {"value":[{"id":"p1","roles":["write"],
                  "grantedTo":{"user":{"id":"u1","displayName":"Ada"}},
                  "grantedToIdentities":[{"user":{"id":"u2"}}]}]}
                """);
        ListItem item = stub.getListItem(GetListItemRequest.newBuilder().setSiteId(SITE).setListId("L1")
                .setItemId("3").setInclude(ListItemInclude.newBuilder().setDriveItem(true)
                        .setPermissions(true)).build()).getItem();
        assertThat(fake.only(route).query())
                .isEqualTo("$select=" + GraphPaths.LIST_ITEM_SELECT + "&$expand=fields,driveItem");
        assertThat(item.getDriveItem().getDriveId()).isEqualTo("d1");
        Permission permission = item.getPermissions(0);
        assertThat(permission.getGrantedTo().getUser().getDisplayName()).isEqualTo("Ada");
        assertThat(permission.getGrantedToIdentities(0).getUser().getId()).isEqualTo("u2");
        assertThat(permission.getTarget().getListItem())
                .isEqualTo(ListItemRef.newBuilder().setSiteId(SITE).setListId("L1").setItemId("3").build());
    }

    @Test
    void deltaListItemsEmitsDeletesAndRereadsRecordsWithoutFields() {
        String base = "/sites/" + SITE + "/lists/L1/items";
        String d1 = fake.baseUrl() + base + "/delta?token=1";
        fake.stub(base + "/delta", MicrosoftFixtures.deltaJson(d1,
                "{\"id\":\"1\",\"fields\":{\"Title\":\"A\"}}",
                "{\"id\":\"2\"}",
                "{\"id\":\"3\",\"deleted\":{\"state\":\"deleted\"},\"contentType\":{\"id\":\"0x01\"}}"));
        fake.stub(base + "/2", "{\"id\":\"2\",\"fields\":{\"Title\":\"B\"}}");
        List<DeltaListItemsResponse> events = drain(stub.deltaListItems(DeltaListItemsRequest.newBuilder()
                .setSiteId(SITE).setListId("L1").build()));
        assertThat(fake.only(base + "/delta").query())
                .isEqualTo("$select=" + GraphPaths.LIST_ITEM_SELECT + "&$expand=fields");
        assertThat(column(events.get(1).getItem(), "Title").getStringValue()).isEqualTo("B");
        assertThat(events.get(2).getItem().getDeleted().getState()).isEqualTo("deleted");
        assertThat(fake.requestsTo(base + "/3")).isEmpty();
        assertThat(events.get(3).getDeltaLink()).isEqualTo(d1);

        assertThat(code(() -> drain(stub.deltaListItems(DeltaListItemsRequest.newBuilder()
                .setSiteId(SITE).setListId("L1")
                .setDeltaLink(fake.baseUrl() + "/sites/" + SITE + "/lists/L9/items/delta?token=1")
                .build())))).isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    @Test
    void listColumnsOfAContentTypeAreTyped() {
        String route = "/sites/" + SITE + "/contentTypes/0x0101/columns";
        fake.stub(route, """
                {"value":[
                  {"id":"c1","name":"Status","displayName":"Status","type":"choice","indexed":true,
                   "choice":{"choices":["Active","Done"],"displayAs":"dropDownMenu"},
                   "validation":{"formula":"=1","descriptions":[{"displayName":"x","languageTag":"en"}]}},
                  {"id":"c2","name":"Cost","number":{"maximum":10,"decimalPlaces":"two"}},
                  {"id":"c3","name":"Ref","lookup":{"listId":"L2","columnName":"Title"},
                   "sourceContentType":{"id":"0x01","name":"Item"}},
                  {"id":"c4","name":"Flag","boolean":{}, "type":"futureThing"}]}
                """);
        List<ColumnDefinition> columns = drain(stub.listColumns(ListColumnsRequest.newBuilder()
                .setSiteId(SITE).setContentTypeId("0x0101").build())).stream().map(r -> r.getColumn()).toList();
        assertThat(fake.only(route).query()).isEmpty();
        assertThat(columns.get(0).getType()).isEqualTo(ColumnType.COLUMN_TYPE_CHOICE);
        assertThat(columns.get(0).getChoice().getChoicesList()).containsExactly("Active", "Done");
        assertThat(columns.get(0).getValidation().getDescriptions(0).getLanguageTag()).isEqualTo("en");
        assertThat(columns.get(0).getContentTypeId()).isEqualTo("0x0101");
        assertThat(columns.get(1).getNumber().getMaximum()).isEqualTo(10);
        assertThat(columns.get(1).getNumber().hasMinimum()).isFalse();
        assertThat(columns.get(2).getLookup().getListId()).isEqualTo("L2");
        assertThat(columns.get(2).getSourceContentType().getName()).isEqualTo("Item");
        assertThat(columns.get(3).hasBoolean()).isTrue();
        assertThat(columns.get(3).getType()).isEqualTo(ColumnType.COLUMN_TYPE_UNKNOWN_FUTURE_VALUE);
    }

    // ------------------------------------------------------------ permissions

    @Test
    void listPermissionsOfASiteAndALink() {
        fake.stub("/sites/" + SITE + "/permissions", """
                {"value":[{"id":"a1","roles":["write"],
                  "grantedToIdentitiesV2":[{"application":{"id":"app-1","displayName":"Indexer"}}]}]}
                """);
        fake.stub("/drives/d1/items/x1/permissions", """
                {"value":[{"id":"l1","roles":["read"],"shareId":"s!1","hasPassword":true,
                  "link":{"type":"view","scope":"anonymous","webUrl":"https://1drv.ms/x",
                    "preventsDownload":true},
                  "expirationDateTime":"2030-01-01T00:00:00Z",
                  "grantedToIdentitiesV2":[{"siteGroup":{"id":"7","displayName":"Visitors",
                    "loginName":"c:0"}},{"sharePointGroup":{"id":"g","title":"Owners"}}],
                  "invitation":{"email":"x@y.com","signInRequired":true}}]}
                """);
        Permission site = drain(stub.listPermissions(ListPermissionsRequest.newBuilder()
                .setTarget(PermissionTarget.newBuilder().setSiteId(SITE)).build())).get(0).getPermission();
        assertThat(site.getGrantedToIdentities(0).getApplication().getDisplayName()).isEqualTo("Indexer");
        assertThat(site.getTarget().getSiteId()).isEqualTo(SITE);
        Permission link = drain(stub.listPermissions(ListPermissionsRequest.newBuilder()
                .setTarget(PermissionTarget.newBuilder().setDriveItem(
                        DriveItemRef.newBuilder().setDriveId("d1").setItemId("x1"))).build()))
                .get(0).getPermission();
        assertThat(link.getLink().getScope()).isEqualTo("anonymous");
        assertThat(link.getLink().getPreventsDownload()).isTrue();
        assertThat(link.getHasPassword()).isTrue();
        assertThat(link.hasExpiresAt()).isTrue();
        assertThat(link.getGrantedToIdentities(0).getSiteGroup().getLoginName()).isEqualTo("c:0");
        assertThat(link.getGrantedToIdentities(1).getSharePointGroup().getTitle()).isEqualTo("Owners");
        assertThat(link.getInvitation().getSignInRequired()).isTrue();
        assertThat(MicrosoftEntityIds.permission(link)).isEqualTo("drive_item/d1/x1/l1");
        assertThat(code(() -> drain(stub.listPermissions(ListPermissionsRequest.getDefaultInstance()))))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
    }

    // ------------------------------------------------------------ sync

    @Test
    void syncSweepsSitesDrivesAndListsThenResumesFromItsCursor() {
        String drives = "/sites/" + SITE + "/drives";
        String lists = "/sites/" + SITE + "/lists";
        String driveLink1 = fake.baseUrl() + "/drives/d1/root/delta?token=a";
        String listLink1 = fake.baseUrl() + lists + "/L1/items/delta?token=a";
        fake.stub("/sites/" + SITE, "{\"id\":\"" + SITE + "\",\"name\":\"finance\"}");
        fake.stub(drives, """
                {"value":[{"id":"d1","name":"Docs","driveType":"documentLibrary"},
                          {"id":"d9","name":"Site Assets","driveType":"documentLibrary","system":{}}]}
                """);
        fake.stub("/drives/d1/root/delta", MicrosoftFixtures.deltaJson(driveLink1,
                MicrosoftFixtures.fileJson("x1", "a.txt", "d1")));
        fake.stub(lists, """
                {"value":[{"id":"L1","name":"Tasks","list":{"template":"genericList"}},
                          {"id":"L0","name":"Docs","list":{"template":"documentLibrary"}}]}
                """);
        fake.stub(lists + "/L1/items/delta", MicrosoftFixtures.deltaJson(listLink1,
                "{\"id\":\"1\",\"fields\":{\"Title\":\"T\"}}"));

        List<SyncResponse> first = drain(stub.sync(SyncRequest.newBuilder().addSiteIds(SITE)
                .setIncludeLists(true).build()));
        assertThat(fake.requestsTo("/drives/d9/root/delta")).isEmpty();
        assertThat(fake.requestsTo(lists + "/L0/items/delta")).isEmpty();
        assertThat(first.stream().filter(SyncResponse::hasChange).map(r -> r.getChange().getEntity()
                .getEntityId()).toList()).containsExactly(SITE, "d1", "x1", "L1", "L1/1");
        assertThat(first.stream().filter(SyncResponse::hasSnapshot).map(r -> r.getSnapshot().getSource())
                .toList()).containsExactly(ChangeSource.CHANGE_SOURCE_CRAWL, ChangeSource.CHANGE_SOURCE_CRAWL);
        String resume = first.get(first.size() - 1).getResumeCursor();
        SyncCursor cursor = SyncCursors.decode(resume, fake.baseUrl());
        assertThat(cursor.getDriveDeltaLinksMap()).containsEntry("d1", driveLink1);
        assertThat(cursor.getListDeltaLinksMap()).containsEntry(SITE + "/L1", listLink1);

        fake.stub("/drives/d1/root/delta?token=a", MicrosoftFixtures.deltaJson(driveLink1 + "b",
                MicrosoftFixtures.deletedJson("x1", "d1")));
        fake.stub(lists + "/L1/items/delta?token=a", MicrosoftFixtures.deltaJson(listLink1 + "b",
                "{\"id\":\"1\",\"deleted\":{\"state\":\"deleted\"}}"));
        List<SyncResponse> second = drain(stub.sync(SyncRequest.newBuilder().addSiteIds(SITE)
                .setIncludeLists(true).setSinceCursor(resume).build()));
        List<ai.pipestream.microsoft.v1.MicrosoftChange> deletes = second.stream()
                .filter(SyncResponse::hasChange).map(SyncResponse::getChange)
                .filter(c -> c.getOperation() == ChangeOperation.CHANGE_OPERATION_DELETE).toList();
        assertThat(deletes).extracting(c -> c.getEntity().getEntityId()).containsExactly("x1", "L1/1");
        assertThat(second.stream().filter(SyncResponse::hasSnapshot).map(r -> r.getSnapshot().getSource())
                .toList()).containsExactly(ChangeSource.CHANGE_SOURCE_DELTA, ChangeSource.CHANGE_SOURCE_DELTA);
        assertThat(SyncCursors.decode(second.get(second.size() - 1).getResumeCursor(), fake.baseUrl())
                .getDriveDeltaLinksMap()).containsEntry("d1", driveLink1 + "b");
    }

    @Test
    void syncRejectsACursorPointingOffTheGraphHost() {
        String forged = SyncCursors.encode(SyncCursor.newBuilder()
                .putDriveDeltaLinks("d1", "https://attacker.example/drives/d1/root/delta?token=x").build());
        int before = fake.requests().size();
        assertThat(code(() -> drain(stub.sync(SyncRequest.newBuilder().setSinceCursor(forged).build()))))
                .isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(code(() -> drain(stub.sync(SyncRequest.newBuilder().setSinceCursor("not-a-cursor!")
                .build())))).isEqualTo(Status.Code.INVALID_ARGUMENT);
        assertThat(fake.requests()).hasSize(before);
    }

    private static ListColumn column(ListItem item, String name) {
        return item.getFieldsList().stream().filter(c -> c.getName().equals(name)).findFirst()
                .orElseThrow(() -> new AssertionError("no column " + name + " in " + item.getFieldsList()));
    }

    @Test
    void unsupportedSelectFallsBackToGraphDefaults() {
        fake.stub("/drives/d1/items/x1?$select=" + GraphPaths.DRIVE_ITEM_SELECT,
                FakeGraphServer.error(400, "BadRequest", Map.of()));
        fake.stub("/drives/d1/items/x1", MicrosoftFixtures.fileJson("x1", "a.txt", "d1"));
        DriveItem item = stub.getItem(GetItemRequest.newBuilder().setDriveId("d1").setItemId("x1")
                .setInclude(DriveItemInclude.getDefaultInstance()).build()).getItem();
        assertThat(item.getName()).isEqualTo("a.txt");
        assertThat(fake.requestsTo("/drives/d1/items/x1")).extracting(r -> r.query())
                .containsExactly("$select=" + GraphPaths.DRIVE_ITEM_SELECT, "");
        assertThatThrownBy(() -> GraphPaths.segment("..", "id")).isInstanceOf(IllegalArgumentException.class);
    }
}
