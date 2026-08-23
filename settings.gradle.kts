plugins {
    id("org.gradle.toolchains.foojay-resolver-convention") version "1.0.0"
}

rootProject.name = "grpc-confluence"

include("grpc-confluence-api")
include("grpc-confluence-service")
include("grpc-microsoft-api")
include("grpc-microsoft-service")
include("grpc-microsoft-connector")
include("grpc-connect")
include("grpc-sync-api")
include("grpc-sync-service")
include("grpc-mcp")

dependencyResolutionManagement {
    repositories {
        mavenCentral()
    }
}
