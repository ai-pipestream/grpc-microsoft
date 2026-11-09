import com.google.protobuf.gradle.id

plugins {
    application
    alias(libs.plugins.protobuf)
}

dependencies {
    implementation(project(":grpc-microsoft-api"))
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.protobuf)
    implementation(libs.grpc.stub)
    implementation(libs.grpc.services)
    implementation(libs.protobuf.java)
    implementation(libs.jackson.databind)
    compileOnly(libs.tomcat.annotations)
    runtimeOnly(libs.log4j.core)

    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj)
    testRuntimeOnly(libs.junit.launcher)
}

protobuf {
    protoc {
        artifact = "com.google.protobuf:protoc:${libs.versions.protobuf.get()}"
    }
    plugins {
        id("grpc") {
            artifact = "io.grpc:protoc-gen-grpc-java:${libs.versions.grpc.get()}"
        }
    }
    generateProtoTasks {
        all().forEach { task ->
            task.plugins {
                id("grpc")
            }
        }
    }
}

application {
    mainClass = "ai.pipestream.microsoft.connector.ConnectorServer"
}
