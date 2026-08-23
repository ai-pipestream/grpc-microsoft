plugins {
    application
}

dependencies {
    implementation(project(":grpc-microsoft-api"))
    implementation(project(":grpc-sync-api"))
    implementation(project(":grpc-okf"))
    implementation(project(":grpc-output-spi"))
    runtimeOnly(project(":grpc-output-filesystem"))
    runtimeOnly(project(":grpc-output-s3"))
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.services)
    implementation(libs.jackson.databind)
    implementation(libs.kafka.clients)
    runtimeOnly(libs.log4j.core)

    testImplementation(project(":grpc-sync-service"))
    testImplementation(project(":grpc-output-filesystem"))
    testImplementation(project(":grpc-output-s3"))
    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

application {
    mainClass = "ai.pipestream.microsoft.MicrosoftServer"
}

tasks.named<Test>("test") {
    exclude("**/MicrosoftLiveSmokeIT.class")
}

tasks.register<Test>("liveSmokeTest") {
    description = "One cheap read against the real Microsoft Graph tenant; needs MICROSOFT_TENANT_ID + MICROSOFT_CLIENT_ID + MICROSOFT_CLIENT_SECRET, skips otherwise"
    group = "verification"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    filter {
        includeTestsMatching("ai.pipestream.microsoft.MicrosoftLiveSmokeIT")
    }
}
