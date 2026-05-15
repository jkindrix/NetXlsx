#!/usr/bin/env bash
# Local + CI build entry point (Unix). See design §9.3 / decision S18.
# Mirrors build/build.ps1; both call the same MSBuild targets.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

CONFIGURATION="${CONFIGURATION:-Release}"
TARGET="${1:-all}"

echo "==> NetXlsx build ($CONFIGURATION) target=$TARGET"

case "$TARGET" in
  restore)
    dotnet restore NetXlsx.sln
    ;;
  build)
    dotnet build NetXlsx.sln -c "$CONFIGURATION" --no-restore
    ;;
  test)
    dotnet test NetXlsx.sln -c "$CONFIGURATION" --no-build \
      --logger "trx;LogFileName=test-results.trx" \
      --collect:"XPlat Code Coverage"
    ;;
  pack)
    dotnet pack src/NetXlsx/NetXlsx.csproj -c "$CONFIGURATION" --no-build \
      -o artifacts/nupkg
    ;;
  bench)
    dotnet run --project benchmarks/NetXlsx.Benchmarks -c "$CONFIGURATION" -- "$@"
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
