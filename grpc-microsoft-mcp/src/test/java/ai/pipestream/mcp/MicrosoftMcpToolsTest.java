package ai.pipestream.mcp;

import ai.pipestream.confluence.v1.ConfluenceServiceGrpc;
import ai.pipestream.confluence.v1.GetPageRequest;
import ai.pipestream.confluence.v1.GetPageResponse;
import ai.pipestream.confluence.v1.ListSpacesRequest;
import ai.pipestream.confluence.v1.ListSpacesResponse;
import ai.pipestream.confluence.v1.Page;
import ai.pipestream.confluence.v1.Space;
import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.Server;
import io.grpc.inprocess.InProcessChannelBuilder;
import io.grpc.inprocess.InProcessServerBuilder;
import io.grpc.stub.StreamObserver;
import io.modelcontextprotocol.client.McpClient;
import io.modelcontextprotocol.client.McpSyncClient;
import io.modelcontextprotocol.client.transport.HttpClientStreamableHttpTransport;
import io.modelcontextprotocol.spec.McpSchema;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class ConnectorMcpToolsTest {

    private Server grpc;
    private McpHttpServer mcp;

    @AfterEach
    void stop() {
        if (mcp != null) {
            mcp.close();
        }
        if (grpc != null) {
            grpc.shutdownNow();
        }
    }

    @Test
    void confluenceAndMicrosoftToolsCallStreamingStubs() throws Exception {
        String name = InProcessServerBuilder.generateName();
        grpc = InProcessServerBuilder.forName(name)
                .directExecutor()
                .addService(new ConfluenceServiceGrpc.ConfluenceServiceImplBase() {
                    @Override
                    public void listSpaces(ListSpacesRequest request,
                            StreamObserver<ListSpacesResponse> observer) {
                        observer.onNext(ListSpacesResponse.newBuilder()
                                .addSpaces(Space.newBuilder()
                                        .setId("100")
                                        .setKey("ENG")
                                        .setName("Engineering"))
                                .build());
                        observer.onCompleted();
                    }

                    @Override
                    public void getPage(GetPageRequest request,
                            StreamObserver<GetPageResponse> observer) {
                        observer.onNext(GetPageResponse.newBuilder()
                                .setPage(Page.newBuilder()
                                        .setId(request.getId())
                                        .setSpaceId("100")
                                        .setTitle("Design")
                                        .setWebUrl("https://example/wiki/pages/" + request.getId()))
                                .build());
                        observer.onCompleted();
                    }
                })
                .addService(new MicrosoftServiceGrpc.MicrosoftServiceImplBase() {
                    @Override
                    public void getMe(GetMeRequest request, StreamObserver<GetMeResponse> observer) {
                        observer.onNext(GetMeResponse.newBuilder()
                                .setUser(GraphUser.newBuilder()
                                        .setId("user-1")
                                        .setDisplayName("Bot")
                                        .setUserPrincipalName("bot@contoso.com"))
                                .build());
                        observer.onCompleted();
                    }
                })
                .build()
                .start();
        var confluence = ConfluenceServiceGrpc.newBlockingStub(
                InProcessChannelBuilder.forName(name).directExecutor().build());
        var microsoft = MicrosoftServiceGrpc.newBlockingStub(
                InProcessChannelBuilder.forName(name).directExecutor().build());
        List<io.modelcontextprotocol.server.McpServerFeatures.SyncToolSpecification> tools =
                new ArrayList<>();
        tools.addAll(ConfluenceMcpTools.of(confluence));
        tools.addAll(MicrosoftMcpTools.of(microsoft));
        mcp = McpHttpServer.start(0, "test-mcp", tools);

        HttpClientStreamableHttpTransport transport = HttpClientStreamableHttpTransport
                .builder("http://127.0.0.1:" + mcp.port())
                .endpoint(McpHttpServer.ENDPOINT)
                .build();
        try (McpSyncClient client = McpClient.sync(transport)
                .requestTimeout(Duration.ofSeconds(15))
                .build()) {
            client.initialize();
            assertThat(client.listTools().tools())
                    .extracting(McpSchema.Tool::name)
                    .contains("confluence_list_spaces", "confluence_get_page",
                            "microsoft_get_me", "microsoft_sync");

            McpSchema.CallToolResult spaces = client.callTool(McpSchema.CallToolRequest.builder()
                    .name("confluence_list_spaces")
                    .arguments(Map.of("limit", 1))
                    .build());
            assertThat(spaces.isError()).isFalse();
            assertThat(spaces.content().get(0).toString()).contains("ENG");

            McpSchema.CallToolResult page = client.callTool(McpSchema.CallToolRequest.builder()
                    .name("confluence_get_page")
                    .arguments(Map.of("id", "200"))
                    .build());
            assertThat(page.isError()).isFalse();
            assertThat(page.content().get(0).toString()).contains("Design");

            McpSchema.CallToolResult me = client.callTool(McpSchema.CallToolRequest.builder()
                    .name("microsoft_get_me")
                    .arguments(Map.of())
                    .build());
            assertThat(me.isError()).isFalse();
            assertThat(me.content().get(0).toString()).contains("bot@contoso.com");
        }
    }
}
