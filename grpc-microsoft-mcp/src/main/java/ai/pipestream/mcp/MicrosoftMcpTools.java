package ai.pipestream.mcp;

import ai.pipestream.microsoft.v1.GetMeRequest;
import ai.pipestream.microsoft.v1.MicrosoftServiceGrpc;
import ai.pipestream.microsoft.v1.SyncRequest;
import ai.pipestream.microsoft.v1.SyncResponse;
import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.server.McpSyncServerExchange;
import io.modelcontextprotocol.spec.McpSchema;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;

final class MicrosoftMcpTools {

    private MicrosoftMcpTools() {
    }

    static List<McpServerFeatures.SyncToolSpecification> of(
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub) {
        return List.of(me(stub), sync(stub));
    }

    private static McpServerFeatures.SyncToolSpecification me(
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub) {
        return tool("microsoft_get_me", "The signed-in Graph user (cheap connectivity probe)",
                McpJson.objectSchema(Map.of(), List.of()),
                (exchange, request) -> {
                    var user = stub.getMe(GetMeRequest.getDefaultInstance()).getUser();
                    return text(McpJson.write(Map.of(
                            "id", user.getId(),
                            "displayName", user.getDisplayName(),
                            "userPrincipalName", user.getUserPrincipalName())));
                });
    }

    private static McpServerFeatures.SyncToolSpecification sync(
            MicrosoftServiceGrpc.MicrosoftServiceBlockingStub stub) {
        return tool("microsoft_sync",
                "Run one streaming Microsoft Graph Sync pass of drive items",
                McpJson.objectSchema(Map.of(
                        "includeContent", McpJson.boolProp("Inline file bytes when size-capped"),
                        "folderPath", McpJson.stringProp("Folder to start from"),
                        "limit", McpJson.intProp("Max change events to return"),
                        "connectionId", McpJson.stringProp("Catalog connection; empty = default")),
                        List.of()),
                (exchange, request) -> {
                    int limit = McpJson.argInt(request.arguments(), "limit", 50);
                    Iterator<SyncResponse> stream = stub.sync(SyncRequest.newBuilder()
                            .setIncludeContent(McpJson.argBool(request.arguments(),
                                    "includeContent", false))
                            .setFolderPath(McpJson.argString(request.arguments(), "folderPath", "/"))
                            .setConnectionId(McpJson.argString(request.arguments(),
                                    "connectionId", ""))
                            .build());
                    List<Map<String, String>> sample = new ArrayList<>();
                    int changes = 0;
                    int snapshots = 0;
                    String resume = "";
                    while (stream.hasNext()) {
                        SyncResponse event = stream.next();
                        if (event.hasChange()) {
                            changes++;
                            if (sample.size() < limit) {
                                sample.add(Map.of(
                                        "changeId", event.getChange().getChangeId(),
                                        "kind", event.getChange().getEntity().getEntityCase().name(),
                                        "entityId", event.getChange().getEntity().getEntityId()));
                            }
                        } else if (event.hasSnapshot()) {
                            snapshots++;
                        } else if (event.getEventCase() == SyncResponse.EventCase.RESUME_CURSOR) {
                            resume = event.getResumeCursor();
                        }
                    }
                    return text(McpJson.write(Map.of(
                            "changes", changes,
                            "snapshots", snapshots,
                            "resumeCursor", resume,
                            "sample", sample)));
                });
    }

    private static McpServerFeatures.SyncToolSpecification tool(String name, String description,
            Map<String, Object> schema,
            BiFunction<McpSyncServerExchange, McpSchema.CallToolRequest, McpSchema.CallToolResult> handler) {
        return McpServerFeatures.SyncToolSpecification.builder()
                .tool(McpSchema.Tool.builder(name, schema).description(description).build())
                .callHandler(handler)
                .build();
    }

    private static McpSchema.CallToolResult text(String body) {
        return McpSchema.CallToolResult.builder()
                .content(List.of(McpSchema.TextContent.builder(body).build()))
                .build();
    }
}
