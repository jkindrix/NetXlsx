// NetXlsx — TeamCity Kotlin DSL pipeline
// See design §9.3 / decision S17. This is the source of truth for CI on the
// TeamCity server; the file is loaded via "Versioned Settings".
//
// STATUS: scaffold placeholder. Wire to a real TeamCity project once the
// github.com/jkindrix (placeholder in Directory.Build.props) is filled in.

import jetbrains.buildServer.configs.kotlin.*
import jetbrains.buildServer.configs.kotlin.buildSteps.script
import jetbrains.buildServer.configs.kotlin.triggers.vcs

version = "2024.03"

project {
    description = "NetXlsx — idiomatic C# facade over NPOI for .xlsx authoring/reading"

    buildType(Build)
    buildType(Pack)
    buildType(Publish)
}

object Build : BuildType({
    name = "Build & Test"
    description = "Restore, build, run xUnit tests including round-trip preservation, " +
                  "concurrent-mutation, use-after-dispose, A1 parser, and golden-file " +
                  "fixtures. Public-API snapshot check enforced."

    artifactRules = """
        **/TestResults/**/*.trx => test-results
        **/TestResults/**/coverage.cobertura.xml => coverage
    """.trimIndent()

    steps {
        script {
            name = "build.sh all"
            scriptContent = "build/build.sh all"
        }
    }

    triggers {
        vcs {}
    }

    requirements {
        // TODO: pin to an agent pool that has libgdiplus + a fallback font
        // installed (decision I3). Without that, the headless-Linux
        // AutoSize test will fail.
    }
})

object Pack : BuildType({
    name = "Pack"
    description = "Produce .nupkg + .snupkg from src/NetXlsx. Versioned via MinVer from git tags."
    dependencies {
        snapshot(Build) {}
        artifacts(Build) { artifactRules = "+:**/*.dll" }
    }
    steps {
        script {
            name = "build.sh pack"
            scriptContent = "build/build.sh pack"
        }
    }
    artifactRules = "artifacts/nupkg/*.nupkg => nupkg\nartifacts/nupkg/*.snupkg => snupkg"
})

object Publish : BuildType({
    name = "Publish to internal feed"
    description = "Push .nupkg + .snupkg to the nuget.org feed. Runs only on tag pushes (vX.Y.Z)."
    dependencies {
        snapshot(Pack) {}
        artifacts(Pack) { artifactRules = "+:nupkg/*.nupkg\n+:snupkg/*.snupkg" }
    }
    steps {
        script {
            name = "Push to internal feed"
            scriptContent = """
                # TODO: replace placeholders once internal feed URL and credentials are known.
                # dotnet nuget push 'nupkg/*.nupkg' \
                #   --source '%env.NUGET_FEED%' \
                #   --api-key '%env.NUGET_API_KEY%' \
                #   --skip-duplicate
                echo 'Publish step disabled — fill in feed URL/credentials in TeamCity project parameters'
                exit 0
            """.trimIndent()
        }
    }
    triggers {
        // TODO: enable a VCS trigger filtered to tag refs matching v[0-9]+\.[0-9]+\.[0-9]+
    }
})
