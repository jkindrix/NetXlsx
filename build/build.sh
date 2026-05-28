#!/usr/bin/env bash
# Local + CI build entry point (Unix). See design §9.3 / decision S18.
# Mirrors build/build.ps1; both call the same MSBuild targets.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

# Prefer a user-level .NET SDK install (~/.dotnet) over the system one
# when present. Lets local developers run newer SDKs without sudo.
if [[ -x "$HOME/.dotnet/dotnet" ]]; then
    export DOTNET_ROOT="$HOME/.dotnet"
    export PATH="$HOME/.dotnet:$PATH"
fi

CONFIGURATION="${CONFIGURATION:-Release}"
TARGET="${1:-all}"
# Shift past the verb so $@ contains only forwarded args (bench filter, etc.)
if [[ $# -gt 0 ]]; then shift || true; fi

echo "==> NetXlsx build ($CONFIGURATION) target=$TARGET"

case "$TARGET" in
  restore)
    dotnet restore NetXlsx.sln
    ;;
  build)
    dotnet build NetXlsx.sln -c "$CONFIGURATION" --no-restore
    ;;
  test)
    # Excludes Category=HeadlessNoFonts to match CI's main test job. Those
    # tests assert font-absent behavior (e.g. AutoSize throwing) and only
    # pass in CI's dedicated fonts-purged job, so they spuriously fail on a
    # workstation with fonts installed.
    dotnet test NetXlsx.sln -c "$CONFIGURATION" --no-build \
      --logger "trx;LogFileName=test-results.trx" \
      --collect:"XPlat Code Coverage" \
      --filter "Category!=HeadlessNoFonts"
    ;;
  pack)
    dotnet pack src/NetXlsx/NetXlsx.csproj -c "$CONFIGURATION" --no-build \
      -o artifacts/nupkg
    ;;
  bench)
    dotnet run --project benchmarks/NetXlsx.Benchmarks -c "$CONFIGURATION" -- "$@"
    # $@ no longer contains the "bench" verb (shifted above), so it forwards
    # only the user's filter/options to BenchmarkDotNet.
    ;;
  all)
    "$0" restore
    "$0" build
    "$0" test
    "$0" pack
    ;;
  *)
    echo "Unknown target: $TARGET" >&2
    echo "Usage: $0 [restore|build|test|pack|bench|all]" >&2
    exit 2
    ;;
esac
