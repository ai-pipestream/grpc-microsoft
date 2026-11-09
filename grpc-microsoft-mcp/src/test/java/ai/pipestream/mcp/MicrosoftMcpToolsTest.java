package ai.pipestream.mcp;

import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.GetMeResponse;
import ai.pipestream.microsoft.v1.GraphUser;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.ManagedChannel;
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
import java.util.Map;

import static org.assertj.core.api.Assertions.assertThat;

class MicrosoftMcpToolsTest {

    private Server grpc;
    private ManagedChannel channel;
    private McpHttpServer mcp;

    @AfterEach
    void stop() {
        if (mcp != null) {
            mcp.close();
        }
        if (channel != null) {
            channel.shutdownNow();
        }
        if (grpc != null) {
            grpc.shutdownNow();
        }
    }

    @Test
    void providerIsRegisteredForTheCombinedServer() {
        assertThat(McpToolProviders.load())
                .filteredOn(provider -> provider instanceof MicrosoftMcpToolProvider)
                .singleElement()
                .satisfies(provider -> {
                    assertThat(provider.name()).isEqualTo("microsoft");
                    assertThat(provider.targetEnvironmentVariable()).isEqualTo("MICROSOFT_GRPC_TARGET");
                    assertThat(provider.defaultTarget()).isEqualTo("localhost:9096");
                });
    }

    @Test
    void microsoftToolsCallStreamingStubs() throws Exception {
        String name = InProcessServerBuilder.generateName();
        grpc = InProcessServerBuilder.forName(name)
                .directExecutor()
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
        channel = InProcessChannelBuilder.forName(name).directExecutor().build();
        mcp = McpHttpServer.start(0, "test-mcp", new MicrosoftMcpToolProvider().tools(channel));

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
                    .contains("microsoft_get_me", "microsoft_sync");

            McpSchema.CallToolResult me = client.callTool(McpSchema.CallToolRequest.builder("microsoft_get_me")
                    .arguments(Map.of())
                    .build());
            assertThat(me.isError()).isFalse();
            assertThat(me.content().get(0).toString()).contains("bot@contoso.com");
        }
    }
}
