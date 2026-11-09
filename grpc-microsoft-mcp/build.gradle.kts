plugins {
    `java-library`
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
