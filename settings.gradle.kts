plugins {
    id("org.gradle.toolchains.foojay-resolver-convention") version "0.10.0"
}

rootProject.name = "grpc-confluence"

include("grpc-confluence-api")
include("grpc-confluence-service")
include("grpc-microsoft-api")
include("grpc-microsoft-service")
include("grpc-microsoft-connector")
include("grpc-connect")

dependencyResolutionManagement {
    repositories {
        mavenCentral()
    }
}
