# NetXlsx

Idiomatic C# facade over [NPOI](https://github.com/nissl-lab/npoi) for creating and reading `.xlsx` spreadsheets.

**Status:** v0.2.0-alpha. First end-to-end round-trip works (create → write cells → save → open → read). Full v1.0 surface (rows, ranges, styling, freeze, AutoSize, typed mapping bodies, streaming write, etc.) lands in subsequent milestones — see [roadmap](docs/roadmap.md).

## Requirements & known limitations

> **⚠ Not compatible with `PublishAot=true` or `PublishTrimmed=true`.**
> The engine (NPOI 2.7.x) uses `System.Xml.Serialization` and `System.Reflection.Emit` paths that AOT and trim cannot satisfy — measured by [spike 4](spikes/results/spike-4-aot-trim.md). A build-time guard ships with the package: setting either property emits MSBuild error `NXLS0100` / `NXLS0101`. The block will lift when NPOI removes those dependencies (track NPOI 3.x).

> **Typed-mapping methods are emitted but not yet executable.** `[Worksheet]`-generated extension methods (`AddRow` / `AddRows` / `ReadRows`) are decorated `[Obsolete(error: true)]` in v0.2.0 — calling them produces a CS0619 compile error pointing to the milestone in which their bodies land. The generator infrastructure, diagnostic catalog (`NXLS0001`–`NXLS0006`), and emitted signatures are all in place; the body wiring follows `ISheet.AppendRow` in a subsequent slice.

> **No streaming, no styling, no rows API yet in v0.2.0.** In-memory write only. Row/column/range/style/freeze/merge surfaces land in subsequent slices per the roadmap. The escape hatch (`workbook.Underlying`, `sheet.Underlying`, `cell.Underlying` — all concrete `XSSF*` types) reaches the full NPOI surface for anything not yet wrapped.

> **Not thread-safe.** NPOI is not thread-safe; this facade does not lock. Concurrent mutation produces undefined behavior. Concurrent-mutation detection per design decision #43 lands in a later slice.

## Documentation

- [Design](docs/design.md) — 52 foundational + 22 implementation decisions, full v1.0 interface sketch, performance targets, behavioral specifications, quality gates.
- [Roadmap](docs/roadmap.md) — binary feature matrix v1.0/v1.1/v2.0/v3.0/Never, per-release DoD, process rules.
- [Implementation notes](docs/implementation-notes.md) — patterns and lessons from the implementation phase (not yet a methodology — see file header).
- [NPOI workarounds](docs/npoi-workarounds.md) — catalog of NPOI quirks the facade routes around (currently empty by design).
- [Pre-impl spikes](spikes/results/) — measured outcomes for style-dedup feasibility, streaming back-pressure, async wrapping cost, and AOT/trim posture.

## Layout

```
src/         NetXlsx, NetXlsx.SourceGen
tests/       NetXlsx.Tests, NetXlsx.GoldenFiles, NetXlsx.PublicApi
benchmarks/  NetXlsx.Benchmarks (BenchmarkDotNet)
samples/     NetXlsx.Cookbook
spikes/      NetXlsx.AotSpike + Spike{1,2,3} harnesses + results/
build/       build.sh / build.ps1 (local + CI entry points)
.teamcity/   pipeline-as-code (Kotlin DSL placeholder)
docs/        design, roadmap, implementation-notes, npoi-workarounds
```

## Build

```sh
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
build/build.sh -- spike-1   # run a specific pre-impl spike
```

PowerShell equivalent: `build/build.ps1`. Both scripts auto-detect a user-level `.NET` install under `~/.dotnet` and prefer it over a system install (lets `net9.0` work on a machine whose system SDK is `net8.0`).

## Scaffold placeholders (filled before first publish)

- `nuget.config` — public NuGet feed URL.
- `Directory.Build.props` `RepositoryUrl` — github.com/jkindrix.
- `CODEOWNERS` — owning team identifiers.
- Source Link package wired (depends on Git host choice).

## What works today (v0.2.0)

```csharp
using NetXlsx;

using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");

sheet["A1"].SetString("Region");
sheet["B1"].SetString("Revenue");

sheet["A2"].SetString("North");
sheet["B2"].SetNumber(1234.56m);

await wb.SaveAsync("sales.xlsx");

// Round-trip
using var read = await Workbook.OpenAsync("sales.xlsx");
var name = read["Sales"]["A2"].GetString();        // "North"
var rev  = read["Sales"]["B2"].GetNumber();        // 1234.56
```

Integer cell coordinates will be **1-indexed** to match Excel's UI when the `sheet[r, c]` indexer lands (`sheet[1, 1]` will be `A1`). Today the slice exposes only the `sheet["A1"]` string form.

## Target API (preview, not yet implemented)

```csharp
sheet["A1"].Value("Region").Bold();
sheet["B1"].Value("Revenue").Bold();

sheet[2, 1].Value("North");                          // row 2, col 1 == A2 (1-indexed)
sheet[2, 2].Value(1234.56m).NumberFormat("$#,##0.00");

sheet.Column("B").Width(18);
sheet.FreezeRows(1);
```

The full target surface is specified in [`docs/design.md §6`](docs/design.md).
