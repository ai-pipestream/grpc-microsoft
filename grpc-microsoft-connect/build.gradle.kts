plugins {
    `java-library`
}

description = "Kafka Connect source for Microsoft Graph drives (MicrosoftChange protobuf bytes)"

dependencies {
    api(libs.integration.connect)
    implementation(project(":grpc-microsoft-api"))
    implementation(project(":grpc-microsoft-service"))
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.stub)
    // Supplied by the Connect worker at runtime.
    compileOnly(libs.kafka.connect.api)
    compileOnly(libs.kafka.clients)

    testImplementation(testFixtures(libs.integration.connect))
    testImplementation(libs.kafka.connect.api)
    testImplementation(libs.kafka.clients)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

tasks.register<Zip>("connectPluginZip") {
    group = "distribution"
    description = "Kafka Connect plugin zip (drop into plugin.path)"
    archiveClassifier.set("plugin")
    into("grpc-microsoft-connect") {
        from(tasks.named("jar"))
        from(configurations.runtimeClasspath)
    }
}

tasks.assemble {
    dependsOn("connectPluginZip")
}
