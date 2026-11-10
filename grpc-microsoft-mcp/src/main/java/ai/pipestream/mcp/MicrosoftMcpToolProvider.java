package ai.pipestream.mcp;

import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import io.grpc.Channel;
import io.modelcontextprotocol.server.McpServerFeatures;

import java.util.List;

/**
 * Contributes {@code microsoft_*} tools over {@code MicrosoftService} to the
 * combined MCP server. Registered in
 * {@code META-INF/services/ai.pipestream.mcp.McpToolProvider}.
 */
public final class MicrosoftMcpToolProvider implements McpToolProvider {

    /** Environment variable for the Microsoft gRPC {@code host:port}. */
    public static final String ENV_TARGET = "MICROSOFT_GRPC_TARGET";
    /** Fallback Microsoft target when {@link #ENV_TARGET} is unset. */
    public static final String DEFAULT_TARGET = "localhost:9096";

    /** ServiceLoader constructor. */
    public MicrosoftMcpToolProvider() {
    }

    @Override
    public String name() {
        return "microsoft";
    }

    @Override
    public String targetEnvironmentVariable() {
        return ENV_TARGET;
    }

    @Override
    public String defaultTarget() {
        return DEFAULT_TARGET;
    }

    @Override
    public List<McpServerFeatures.SyncToolSpecification> tools(Channel channel) {
        return MicrosoftMcpTools.of(MicrosoftServiceGrpc.newBlockingStub(channel));
    }
}
