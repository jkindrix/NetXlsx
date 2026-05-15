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

# Prefer a user-level .NET SDK install ($HOME/.dotnet) over the system one
# when present. Lets local developers run newer SDKs without admin rights.
$UserDotnet = Join-Path $HOME '.dotnet'
$UserDotnetExe = if ($IsWindows -or $env:OS -eq 'Windows_NT') {
    Join-Path $UserDotnet 'dotnet.exe'
} else {
    Join-Path $UserDotnet 'dotnet'
}
if (Test-Path $UserDotnetExe) {
    $env:DOTNET_ROOT = $UserDotnet
    $env:PATH = "$UserDotnet$([IO.Path]::PathSeparator)$env:PATH"
}

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
