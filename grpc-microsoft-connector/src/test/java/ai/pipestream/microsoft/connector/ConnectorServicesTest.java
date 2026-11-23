package ai.pipestream.microsoft.connector;

import io.grpc.ManagedChannel;
import io.grpc.Server;
import io.grpc.Status;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import microsoft.graph.connectors.contracts.grpc.AuthenticationData;
import microsoft.graph.connectors.contracts.grpc.ConnectionManagementServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.ConnectorCrawlerServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.ConnectorInfoServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.ConnectorOAuthServiceGrpc;
import microsoft.graph.connectors.contracts.grpc.CrawlCheckpoint;
import microsoft.graph.connectors.contracts.grpc.CrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.CustomConfiguration;
import microsoft.graph.connectors.contracts.grpc.GetBasicConnectorInfoRequest;
import microsoft.graph.connectors.contracts.grpc.GetCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.GetDataSourceSchemaRequest;
import microsoft.graph.connectors.contracts.grpc.GetIncrementalCrawlStreamRequest;
import microsoft.graph.connectors.contracts.grpc.HealthCheckRequest;
import microsoft.graph.connectors.contracts.grpc.IncrementalCrawlStreamBit;
import microsoft.graph.connectors.contracts.grpc.OperationResult;
import microsoft.graph.connectors.contracts.grpc.RefreshAccessTokenRequest;
import microsoft.graph.connectors.contracts.grpc.ValidateAuthenticationRequest;
import microsoft.graph.connectors.contracts.grpc.ValidateCustomConfigurationRequest;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.net.InetSocketAddress;
import java.net.ServerSocket;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * The four GCA services against an in-process MicrosoftService, plus a
 * real Netty listen for the MicrosoftService target the adapter dials.
 */
class ConnectorServicesTest {

    private Server microsoftServer;
    private Server connectorServer;
    private ManagedChannel connectorChannel;
    private FakeMicrosoftService microsoft;
    private ConnectorInfoServiceGrpc.ConnectorInfoServiceBlockingStub info;
    private ConnectionManagementServiceGrpc.ConnectionManagementServiceBlockingStub management;
    private ConnectorCrawlerServiceGrpc.ConnectorCrawlerServiceBlockingStub crawler;
    private ConnectorOAuthServiceGrpc.ConnectorOAuthServiceBlockingStub oauth;
    private String targetJson;

    @BeforeEach
    void startStack() throws Exception {
        microsoft = new FakeMicrosoftService();
        microsoft.changes.add(FakeMicrosoftService.upsertFile("file-1", "notes.txt"));
        // The adapter dials a real host:port, so MicrosoftService also
        // listens on a loopback socket.
        int port = freePort();
        microsoftServer = io.grpc.netty.shaded.io.grpc.netty.NettyServerBuilder
                .forAddress(new InetSocketAddress("127.0.0.1", port))
                .addService(microsoft)
                .build()
                .start();
        targetJson = "{\"target\":\"127.0.0.1:" + port + "\",\"plaintext\":true}";

        String connectorName = InProcessServerBuilder.generateName();
        connectorServer = InProcessServerBuilder.forName(connectorName)
                .directExecutor()
                .addService(new ConnectorInfoServiceImpl())
                .addService(new ConnectionManagementServiceImpl())
                .addService(new ConnectorCrawlerServiceImpl())
                .addService(new ConnectorOAuthServiceImpl())
                .build()
                .start();
        connectorChannel = InProcessChannelBuilder.forName(connectorName).directExecutor().build();
        info = ConnectorInfoServiceGrpc.newBlockingStub(connectorChannel);
        management = ConnectionManagementServiceGrpc.newBlockingStub(connectorChannel);
        crawler = ConnectorCrawlerServiceGrpc.newBlockingStub(connectorChannel);
        oauth = ConnectorOAuthServiceGrpc.newBlockingStub(connectorChannel);
    }

    @AfterEach
    void stopStack() throws InterruptedException {
        if (connectorChannel != null) {
            connectorChannel.shutdownNow();
        }
        if (connectorServer != null) {
            connectorServer.shutdownNow();
            connectorServer.awaitTermination(5, TimeUnit.SECONDS);
        }
        if (microsoftServer != null) {
            microsoftServer.shutdownNow();
            microsoftServer.awaitTermination(5, TimeUnit.SECONDS);
        }
    }

    @Test
    void reportsStableConnectorIdAndHealth() {
        assertThat(info.getBasicConnectorInfo(GetBasicConnectorInfoRequest.getDefaultInstance())
                .getConnectorId()).isEqualTo(ConnectorId.VALUE);
        assertThat(info.healthCheck(HealthCheckRequest.getDefaultInstance())).isNotNull();
    }

    @Test
    void validateAuthenticationCallsGetMe() {
        String hostPort = targetJson.replaceAll(".*\"target\":\"([^\"]+)\".*", "$1");
        var auth = management.validateAuthentication(ValidateAuthenticationRequest.newBuilder()
                .setAuthenticationData(AuthenticationData.newBuilder()
                        .setDatasourceUrl(hostPort))
                .build());
        assertThat(auth.getStatus().getResult()).isEqualTo(OperationResult.Success);
        assertThat(auth.getStatus().getStatusMessage()).contains("bot@contoso.com");

        var withConfig = management.validateCustomConfiguration(
                ValidateCustomConfigurationRequest.newBuilder()
                        .setCustomConfiguration(config())
                        .build());
        assertThat(withConfig.getStatus().getResult()).isEqualTo(OperationResult.Success);
    }

    @Test
    void schemaAndAuthAgainstTheFakeMicrosoftService() {
        var schema = management.getDataSourceSchema(GetDataSourceSchemaRequest.newBuilder()
                .setCustomConfiguration(config())
                .build());
        assertThat(schema.getStatus().getResult()).isEqualTo(OperationResult.Success);
        assertThat(schema.getDataSourceSchema().getPropertyListList())
                .extracting(p -> p.getName())
                .contains(DataSourceSchemas.TITLE, DataSourceSchemas.WEB_URL);

        var auth = ConnectionManagementServiceImpl.parse(config(),
                AuthenticationData.getDefaultInstance());
        assertThat(auth.target()).startsWith("127.0.0.1:");

        // Call GetMe through the real adapter method by stuffing target in JSON
        // and using a request whose custom config is empty... ValidateAuthentication
        // only sees AuthenticationData. Use the crawler to prove the channel.
        List<microsoft.graph.connectors.contracts.grpc.CrawlStreamBit> bits = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(bits::add);

        assertThat(microsoft.syncCalls.get()).isEqualTo(1);
        assertThat(bits).anyMatch(bit -> bit.hasCrawlItem()
                && bit.getCrawlItem().getItemId().equals("drive-1/file-1"));
        assertThat(bits.get(bits.size() - 1).getCrawlProgressMarker().getCustomMarkerData())
                .isEqualTo("cursor-1");
        assertThat(bits).allMatch(bit -> bit.getStatus().getResult() == OperationResult.Success);
    }

    @Test
    void incrementalCrawlResumesFromTheStoredCursor() {
        microsoft.resumeCursor = "cursor-2";

        List<IncrementalCrawlStreamBit> bits = incremental(
                GetIncrementalCrawlStreamRequest.newBuilder()
                        .setCustomConfiguration(config())
                        .setCrawlProgressMarker(CrawlCheckpoint.newBuilder()
                                .setCustomMarkerData("cursor-1"))
                        .build());

        // The stored cursor goes back to Sync, so the pass is a delta pass
        // and carries the deletes since the last one.
        assertThat(microsoft.syncRequests).hasSize(1);
        assertThat(microsoft.syncRequests.get(0).getSinceCursor()).isEqualTo("cursor-1");
        assertThat(bits).allMatch(bit -> bit.getStatus().getResult() == OperationResult.Success);
        // A checkpoint written before the pass ends still names the cursor
        // the pass started from: a crawl resumed from it replays this pass.
        assertThat(bits.subList(0, bits.size() - 1))
                .isNotEmpty()
                .allMatch(IncrementalCrawlStreamBit::hasCrawlItem)
                .extracting(bit -> bit.getCrawlProgressMarker().getCustomMarkerData())
                .containsOnly("cursor-1");
        assertThat(bits.get(bits.size() - 1).getCrawlProgressMarker().getCustomMarkerData())
                .isEqualTo("cursor-2");
    }

    @Test
    void incrementalCrawlWithoutACursorIsAFullEnumeration() {
        List<IncrementalCrawlStreamBit> bits = incremental(
                GetIncrementalCrawlStreamRequest.newBuilder()
                        .setCustomConfiguration(config())
                        .setPreviousCrawlStartTimeInUtc(
                                com.google.protobuf.Timestamp.newBuilder().setSeconds(1_700_000_000))
                        .build());

        assertThat(microsoft.syncRequests.get(0).getSinceCursor()).isEmpty();
        // The previous crawl's start time is not a resume cursor and is
        // never stored as one.
        assertThat(bits.subList(0, bits.size() - 1))
                .extracting(bit -> bit.getCrawlProgressMarker().getCustomMarkerData())
                .containsOnly("");
        assertThat(bits.get(bits.size() - 1).getCrawlProgressMarker().getCustomMarkerData())
                .isEqualTo("cursor-1");
    }

    @Test
    void fullCrawlEnumeratesEverythingAndKeepsTheStoredCursorUntilItEnds() {
        microsoft.resumeCursor = "cursor-2";
        List<CrawlStreamBit> bits = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .setCrawlProgressMarker(CrawlCheckpoint.newBuilder()
                        .setCustomMarkerData("cursor-1"))
                .build()).forEachRemaining(bits::add);

        // A full crawl never asks for a delta pass.
        assertThat(microsoft.syncRequests.get(0).getSinceCursor()).isEmpty();
        assertThat(bits.subList(0, bits.size() - 1))
                .isNotEmpty()
                .extracting(bit -> bit.getCrawlProgressMarker().getCustomMarkerData())
                .containsOnly("cursor-1");
        assertThat(bits.get(bits.size() - 1).getCrawlProgressMarker().getCustomMarkerData())
                .isEqualTo("cursor-2");
    }

    @Test
    void aSyncStreamWithoutItsResumeCursorFailsTheCrawl() {
        microsoft.resumeCursor = null;

        List<IncrementalCrawlStreamBit> incremental = incremental(
                GetIncrementalCrawlStreamRequest.newBuilder()
                        .setCustomConfiguration(config())
                        .setCrawlProgressMarker(CrawlCheckpoint.newBuilder()
                                .setCustomMarkerData("cursor-1"))
                        .build());
        IncrementalCrawlStreamBit lastIncremental = incremental.get(incremental.size() - 1);
        assertThat(lastIncremental.getStatus().getResult())
                .isEqualTo(OperationResult.DatasourceError);
        assertThat(lastIncremental.getStatus().getStatusMessage()).contains("resume cursor");
        // No bit of the truncated pass moves the stored cursor.
        assertThat(incremental)
                .filteredOn(IncrementalCrawlStreamBit::hasCrawlProgressMarker)
                .extracting(bit -> bit.getCrawlProgressMarker().getCustomMarkerData())
                .containsOnly("cursor-1");

        List<CrawlStreamBit> full = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(full::add);
        assertThat(full.get(full.size() - 1).getStatus().getResult())
                .isEqualTo(OperationResult.DatasourceError);
        assertThat(full.get(full.size() - 1).getStatus().getStatusMessage())
                .contains("resume cursor");
    }

    private List<IncrementalCrawlStreamBit> incremental(GetIncrementalCrawlStreamRequest request) {
        List<IncrementalCrawlStreamBit> bits = new ArrayList<>();
        crawler.getIncrementalCrawlStream(request).forEachRemaining(bits::add);
        return bits;
    }

    @Test
    void authFailureIsAuthenticationIssue() {
        microsoft.meError = Status.PERMISSION_DENIED.withDescription("nope");
        // ValidateAuthentication uses default target unless we can pass config.
        // The GCA request has no custom config field; operators put the target
        // in custom configuration on the other two RPCs. Prove mapping via
        // a crawl against a failing Sync instead.
        microsoft.syncError = Status.PERMISSION_DENIED.withDescription("nope");
        List<microsoft.graph.connectors.contracts.grpc.CrawlStreamBit> bits = new ArrayList<>();
        crawler.getCrawlStream(GetCrawlStreamRequest.newBuilder()
                .setCustomConfiguration(config())
                .build()).forEachRemaining(bits::add);
        assertThat(bits).isNotEmpty();
        assertThat(bits.get(0).getStatus().getResult())
                .isEqualTo(OperationResult.AuthenticationIssue);
    }

    @Test
    void oauthRefreshEchoesSuccess() {
        var response = oauth.refreshAccessToken(RefreshAccessTokenRequest.getDefaultInstance());
        assertThat(response.getStatus().getResult()).isEqualTo(OperationResult.Success);
    }

    @Test
    void invalidCustomConfigurationFailsValidation() {
        var response = management.validateCustomConfiguration(
                ValidateCustomConfigurationRequest.newBuilder()
                        .setCustomConfiguration(CustomConfiguration.newBuilder()
                                .setConfiguration("not-json"))
                        .build());
        assertThat(response.getStatus().getResult()).isEqualTo(OperationResult.ValidationFailure);
    }

    private CustomConfiguration config() {
        return CustomConfiguration.newBuilder().setConfiguration(targetJson).build();
    }

    private static int freePort() throws IOException {
        try (ServerSocket socket = new ServerSocket(0)) {
            return socket.getLocalPort();
        }
    }
}
