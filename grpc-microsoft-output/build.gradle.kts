plugins {
    `java-library`
}

dependencies {
    api(project(":grpc-microsoft-api"))
    api(project(":grpc-okf"))
    api(project(":grpc-output-spi"))
    api(libs.protobuf.java)

    testImplementation(testFixtures(project(":grpc-okf")))
    testImplementation(project(":grpc-output-filesystem"))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}
