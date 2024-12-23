plugins {
    kotlin("jvm") version "2.0.20"
    `java-library`
    id("maven-publish")
}

group = "jp.kaiz"
version = "0.0.1"

repositories {
    mavenCentral()
}

dependencies {
    api("org.apache.poi:poi-ooxml:5.2.5")
}

publishing {
    repositories {
        maven {
            name = "GitHubPackages"
            url = uri("https://maven.pkg.github.com/Kai-Z-JP/shachia-poi-dsl")
            credentials {
                username = project.findProperty("gpr.user") as String? ?: System.getenv("USERNAME")
                password = project.findProperty("gpr.key") as String? ?: System.getenv("TOKEN")
            }
        }
    }
    publications {
        create<MavenPublication>("maven") {
            from(components["kotlin"])
        }
    }
}