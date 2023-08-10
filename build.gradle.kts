import groovy.lang.GroovyObject

plugins {
    kotlin("jvm") version "1.5.21"
    id("maven-publish")
}

val poiVersion: String by project
val commonsCsvVersion: String by project

group = "me.kevur"

repositories {
    mavenCentral()
}

dependencies {
    api(kotlin("stdlib-jdk8"))
    implementation("org.apache.poi:poi:${poiVersion}")
    implementation("org.apache.poi:poi-ooxml:${poiVersion}")

    // export api because csv parsing api uses apache commons CSVFormat class
    api("org.apache.commons:commons-csv:${commonsCsvVersion}")
    // CSVs exported from excel can contain BOMs -> we need BOMInputStream to handle this
    implementation("commons-io:commons-io:2.7")

    testImplementation("org.junit.jupiter:junit-jupiter:5.6.2")
}

tasks {
    withType<org.jetbrains.kotlin.gradle.tasks.KotlinCompile> {
        sourceCompatibility = JavaVersion.VERSION_1_8.toString()
        targetCompatibility = JavaVersion.VERSION_1_8.toString()
        kotlinOptions.jvmTarget = "1.8"
    }

    test {
        useJUnitPlatform()
    }
}

publishing {
    publications {
        create<MavenPublication>("lib") {
            artifactId = rootProject.name
            from(components.getByName("java"))
            artifact(tasks.kotlinSourcesJar.get())
        }
    }
}
