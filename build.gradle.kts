import org.jetbrains.compose.desktop.application.dsl.TargetFormat

plugins {
    kotlin("multiplatform")
    id("org.jetbrains.compose")
}




group = "com.example"
version = "1.0-SNAPSHOT"

repositories {
    google()
    mavenCentral()
    maven("https://maven.pkg.jetbrains.space/public/p/compose/dev")
}

kotlin {
    jvm {
        jvmToolchain(11)
        withJava()
    }
    sourceSets {
        val jvmMain by getting {
            dependencies {
                implementation(compose.desktop.currentOs)
                implementation("com.darkrockstudios:mpfilepicker-desktop:1.0.0")
                implementation("io.github.evanrupert:excelkt:1.0.2")
                implementation("org.apache.poi:poi-ooxml:5.0.0")
                implementation("org.apache.poi:poi:5.0.0")
            }
        }
        val jvmTest by getting
    }
}

compose.desktop {
    application {
        mainClass = "MainKt"
        nativeDistributions {
            targetFormats(TargetFormat.Exe, TargetFormat.Dmg, TargetFormat.Msi, TargetFormat.Deb)
            packageName = "demo"
            packageVersion = "1.0.0"
            windows {
                packageVersion = "1.0.0"
                packageName = "demo"
                exePackageVersion = "1.0.0"
                msiPackageVersion = "1.0.0"
            }
        }
    }
}

