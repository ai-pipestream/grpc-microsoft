plugins {
    application
}

description = "Admin service that grants applications access to SharePoint sites (runs with Sites.FullControl.All)"

dependencies {
    implementation(project(":grpc-microsoft-api"))
    implementation(project(":grpc-microsoft-service"))
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.services)
    implementation(libs.jackson.databind)
    runtimeOnly(libs.log4j.core)

    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

application {
    mainClass = "ai.pipestream.microsoft.admin.MicrosoftAdminServer"
}

// The live test hits the real tenant, so it never runs in the default test task.
tasks.named<Test>("test") {
    exclude("**/MicrosoftAdminLiveIT.class")
}

tasks.register<Test>("liveAdminTest") {
    description = "Reads the application grants of MICROSOFT_SITE_ID with the admin identity and dry-runs an ensure; needs MICROSOFT_TENANT_ID + MICROSOFT_ADMIN_CLIENT_ID + a MICROSOFT_ADMIN_CLIENT_CERT_PEM or MICROSOFT_ADMIN_CLIENT_SECRET, skips otherwise"
    group = "verification"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    filter {
        includeTestsMatching("ai.pipestream.microsoft.admin.MicrosoftAdminLiveIT")
    }
    // A live run is never up to date, and the report goes to stdout.
    outputs.upToDateWhen { false }
    testLogging {
        showStandardStreams = true
    }
}
