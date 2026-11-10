plugins {
    application
}

dependencies {
    implementation(project(":grpc-microsoft-api"))
    implementation(libs.integration.sync.api)
    implementation(libs.integration.okf)
    implementation(project(":grpc-microsoft-output"))
    implementation(libs.integration.output.spi)
    runtimeOnly(libs.integration.output.filesystem)
    runtimeOnly(libs.integration.output.s3)
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.services)
    implementation(libs.jackson.databind)
    implementation(libs.kafka.clients)
    runtimeOnly(libs.log4j.core)

    testImplementation(libs.integration.sync.service)
    testImplementation(libs.integration.output.filesystem)
    testImplementation(libs.integration.output.s3)
    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

application {
    mainClass = "ai.pipestream.microsoft.MicrosoftServer"
}

// The live smoke and coverage tests hit the real tenant, so they never run
// in the default test task (and are assumption-gated on credentials as a second guard).
tasks.named<Test>("test") {
    exclude("**/MicrosoftLiveSmokeIT.class")
    exclude("**/MicrosoftLiveCoverageIT.class")
    exclude("**/MicrosoftLiveWriteIT.class")
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

tasks.register<Test>("liveCoverageTest") {
    description = "Calls every MicrosoftService rpc against the real Microsoft Graph tenant and writes build/live-coverage.md; needs MICROSOFT_TENANT_ID + MICROSOFT_CLIENT_ID + MICROSOFT_CLIENT_SECRET, optional MICROSOFT_SITE_ID / MICROSOFT_DRIVE_IDS, skips otherwise"
    group = "verification"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    filter {
        includeTestsMatching("ai.pipestream.microsoft.MicrosoftLiveCoverageIT")
    }
    systemProperty("liveCoverage.report", layout.buildDirectory.file("live-coverage.md").get().asFile.path)
    // A live run is never up to date, and the report goes to stdout too.
    outputs.upToDateWhen { false }
    testLogging {
        showStandardStreams = true
    }
}

tasks.register<Test>("liveWriteTest") {
    description = "Calls every MicrosoftService write rpc against the real Microsoft Graph tenant, inside one new test folder and one new test list that it removes, and writes build/live-write.md; needs MICROSOFT_TENANT_ID + MICROSOFT_CLIENT_ID + MICROSOFT_CLIENT_SECRET + MICROSOFT_SITE_ID and MICROSOFT_LIVE_WRITE=1, optional MICROSOFT_DRIVE_IDS, skips otherwise"
    group = "verification"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    filter {
        includeTestsMatching("ai.pipestream.microsoft.MicrosoftLiveWriteIT")
    }
    systemProperty("liveWrite.report", layout.buildDirectory.file("live-write.md").get().asFile.path)
    // A live run is never up to date, and the report goes to stdout too.
    outputs.upToDateWhen { false }
    testLogging {
        showStandardStreams = true
    }
}
