# Contributing to NetXlsx

Thanks for the interest. This project follows a deliberate design-then-implement loop — that's the practice this codebase exists to embody as much as the API it exposes.

## Before opening a PR

1. **Read [`docs/design.md`](docs/design.md).** It's the long doc but it's the load-bearing one — 52 foundational + 22 implementation-level numbered decisions with rationale columns. New public API surface should be discussable against an existing decision (or motivate a new numbered one).
2. **Check [`docs/roadmap.md`](docs/roadmap.md).** v1.0/1.1/2.0/Never matrix. If your idea is in a `Never` column, see if the rationale still applies before opening — sometimes it does, sometimes it's time to re-examine.
3. **The public-API analyzer enforces the surface.** `Microsoft.CodeAnalysis.PublicApiAnalyzers` is wired: any new public type/member must land alongside an entry in `PublicAPI.Unshipped.txt` or the build fails (`RS0016`). Removed members go via `PublicAPI.Shipped.txt` removal + analyzer release notes. There's a runtime backstop test (`PublicApiSnapshotTests`) that asserts the type list — keep its baseline in sync.

## Code expectations

- **Conventional commits.** Subjects use `feat:`, `fix:`, `docs:`, `chore:`, `refactor:`, `test:`. Bodies prefer "what changed and why" over "how." Slice-level granularity is the norm (see CHANGELOG for examples).
- **Strict warnings.** The project treats warnings as errors. CI enforces. If you suppress a diagnostic, leave a one-line rationale comment.
- **Tests required for behavior changes.** The repo runs 433+ tests/TFM across three projects (unit, golden-file, public-API snapshot). If your change affects behavior, the test count should change too.
- **No AI-tool branding in commits.** Project policy. If your editor or assistant tries to add it, strip it before pushing.

## Pre-impl spikes for risky changes

If your PR touches an area where the cost/benefit is uncertain (perf assumptions, NPOI corner cases, AOT/trim posture, etc.), a measured pre-impl spike is welcome under `spikes/`. The existing spikes are short C# programs with a `RESULT:` writeup in `spikes/results/`. Spikes that contradict a design assumption are *good* — that's exactly what they're for. Update the design doc to record what was learned.

## Running locally

```bash
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
```

PowerShell equivalent: `build/build.ps1`.

### SDK requirements

The project targets `net8.0`, `net9.0`, and `net10.0`. **The .NET 10 SDK is required** to build the repository — `global.json` pins it with `rollForward: latestFeature` (any `10.x` version is fine). The .NET 8 + .NET 9 runtimes are also needed for the test matrix.

Install via your platform's package manager or the official installer at [dotnet.microsoft.com](https://dotnet.microsoft.com/download). On Linux/WSL the `dotnet-install` script with `--channel 10.0` works and drops the SDK into `~/.dotnet/` without root.

Both build scripts (`build/build.sh`, `build/build.ps1`) auto-detect a user-level .NET install under `~/.dotnet` and prefer it over a system install — useful if your system SDK is older than 10.x or if you maintain multiple SDK versions side-by-side.

## Reporting issues

GitHub Issues for bugs and feature requests. Security issues should go via [private security advisories](https://github.com/jkindrix/NetXlsx/security/advisories) per [`SECURITY.md`](SECURITY.md) — not a public issue.

## License

By contributing, you agree your contributions are licensed under the project's [MIT License](LICENSE).
