plugins {
    `java-library`
}

dependencies {
    api(project(":grpc-microsoft-api"))
    api(libs.integration.okf)
    api(libs.integration.output.spi)
    api(libs.protobuf.java)

    testImplementation(testFixtures(libs.integration.okf))
    testImplementation(libs.integration.output.filesystem)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}
