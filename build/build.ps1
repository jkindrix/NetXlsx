<#
.SYNOPSIS
    Local + CI build entry point (Windows / PowerShell). See design §9.3, decision S18.
    Mirrors build/build.sh; both call the same MSBuild targets.
#>
[CmdletBinding()]
param(
    [ValidateSet('restore', 'build', 'test', 'pack', 'bench', 'all')]
    [string]$Target = 'all',
    [string]$Configuration = $(if ($env:CONFIGURATION) { $env:CONFIGURATION } else { 'Release' })
)

$ErrorActionPreference = 'Stop'
$RepoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
Set-Location $RepoRoot

Write-Host "==> NetXlsx build ($Configuration) target=$Target"

function Invoke-Restore { dotnet restore NetXlsx.sln }
function Invoke-Build   { dotnet build   NetXlsx.sln -c $Configuration --no-restore }
function Invoke-Test {
    dotnet test NetXlsx.sln -c $Configuration --no-build `
        --logger "trx;LogFileName=test-results.trx" `
        --collect:"XPlat Code Coverage"
}
function Invoke-Pack {
    dotnet pack src/NetXlsx/NetXlsx.csproj -c $Configuration --no-build `
        -o artifacts/nupkg
}
function Invoke-Bench {
    dotnet run --project benchmarks/NetXlsx.Benchmarks -c $Configuration
}

switch ($Target) {
    'restore' { Invoke-Restore }
    'build'   { Invoke-Build }
    'test'    { Invoke-Test }
    'pack'    { Invoke-Pack }
    'bench'   { Invoke-Bench }
    'all'     { Invoke-Restore; Invoke-Build; Invoke-Test; Invoke-Pack }
}
