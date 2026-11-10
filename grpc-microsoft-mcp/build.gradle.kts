plugins {
    `java-library`
    application
}

description = "Microsoft Graph MCP tools for the combined grpc-mcp server (McpToolProvider)"

dependencies {
    api(libs.integration.mcp)
    implementation(project(":grpc-microsoft-api"))
    compileOnly(libs.tomcat.annotations)

    testImplementation(libs.mcp.core)
    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

// `run` / `installDist` start the combined grpc-mcp server with this
// module's tools on the classpath (plus the core sync-table, connection,
// and app tools). Add other connectors' *-mcp jars for a wider endpoint.
application {
    mainClass = "ai.pipestream.mcp.McpServerMain"
}
