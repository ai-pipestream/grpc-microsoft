import org.gradle.external.javadoc.StandardJavadocDocletOptions

subprojects {
    group = "ai.pipestream.confluence"
    version = "0.1.0-SNAPSHOT"

    plugins.withType<JavaPlugin> {
        extensions.configure<JavaPluginExtension> {
            toolchain {
                languageVersion = JavaLanguageVersion.of(25)
            }
            withJavadocJar()
        }
        tasks.withType<Test>().configureEach {
            useJUnitPlatform()
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
        tasks.named("check") {
            dependsOn("javadoc")
        }
    }
}
