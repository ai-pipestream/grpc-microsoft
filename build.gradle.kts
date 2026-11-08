import org.gradle.external.javadoc.StandardJavadocDocletOptions

// JDK 24+ warns when a library touches sun.misc.Unsafe (protobuf-java's
// UnsafeUtil) or loads native code (grpc-netty-shaded). Both are deliberate
// in those libraries, so the services and tests grant them explicitly.
val runtimeJvmArgs = listOf(
    "--enable-native-access=ALL-UNNAMED",
    "--sun-misc-unsafe-memory-access=allow",
)

subprojects {
    group = "ai.pipestream.confluence"
    version = "0.1.0-SNAPSHOT"

    plugins.withType<JavaPlugin> {
        // grpc-java pulls an older Guava whose AbstractFuture calls the
        // terminally deprecated sun.misc.Unsafe on JDK 25; newer Guava uses
        // VarHandle instead.
        dependencies {
            constraints {
                add("implementation", "com.google.guava:guava:33.8.0-jre")
            }
        }
        extensions.configure<JavaPluginExtension> {
            toolchain {
                languageVersion = JavaLanguageVersion.of(25)
            }
            withJavadocJar()
        }
        tasks.withType<Test>().configureEach {
            useJUnitPlatform()
            jvmArgs(runtimeJvmArgs)
            testLogging {
                events("failed", "skipped")
                showStackTraces = true
            }
        }
        tasks.withType<Javadoc>().configureEach {
            // Generated protobuf / GCA contracts are not our comments.
            exclude("**/ai/pipestream/confluence/v1/**")
            exclude("**/ai/pipestream/microsoft/v1/**")
            exclude("**/ai/pipestream/sync/v1/**")
            exclude("**/microsoft/graph/**")
            options.encoding = "UTF-8"
            (options as StandardJavadocDocletOptions).apply {
                addBooleanOption("Xdoclint:all", true)
                addBooleanOption("Xwerror", true)
                memberLevel = JavadocMemberLevel.PROTECTED
            }
        }
        // Every compiler warning fails the build: deprecation, unchecked,
        // rawtypes and the rest of -Xlint:all.
        tasks.withType<JavaCompile>().configureEach {
            options.encoding = "UTF-8"
            options.compilerArgs.addAll(listOf("-Xlint:all", "-Werror"))
        }
        plugins.withType<ApplicationPlugin> {
            extensions.configure<JavaApplication> {
                applicationDefaultJvmArgs = runtimeJvmArgs
            }
        }
        tasks.named("check") {
            dependsOn("javadoc")
        }
    }
}
