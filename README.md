# NetXlsx

Idiomatic C# facade over [NPOI](https://github.com/nissl-lab/npoi) for creating and reading `.xlsx` spreadsheets.

**Status:** Scaffolded (v0.1.0). No public surface yet — implementation begins after the four pre-implementation spikes (see `docs/roadmap.md`).

## Documentation

- [Design](docs/design.md) — foundational decisions, interface sketch, behavioral specifications, quality gates.
- [Roadmap](docs/roadmap.md) — binary feature matrix, per-release checklists, process rules.
- [NPOI workarounds](docs/npoi-workarounds.md) — catalog of NPOI quirks the facade routes around.

## Layout

```
src/         NetXlsx, NetXlsx.SourceGen
tests/       NetXlsx.Tests, NetXlsx.GoldenFiles, NetXlsx.PublicApi
benchmarks/  NetXlsx.Benchmarks (BenchmarkDotNet)
samples/     NetXlsx.Cookbook
build/       build.sh / build.ps1 (local + CI entry points)
.teamcity/   pipeline-as-code (Kotlin DSL)
docs/        design, roadmap, npoi-workarounds
```

## Build

```sh
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
```

PowerShell equivalent: `build/build.ps1`.

## Scaffold placeholders (TODO)

Filled before the first publish:

- `nuget.config` — public NuGet feed URL.
- `Directory.Build.props` `RepositoryUrl` — github.com/jkindrix.
- `CODEOWNERS` — owning team identifiers.
- Source Link package wired (depends on Git host choice).

## At a glance (target API — not yet implemented)

```csharp
using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");

sheet["A1"].Value("Region").Bold();
sheet["B1"].Value("Revenue").Bold();

sheet[2, 1].Value("North");                          // row 2, col 1 == A2 (1-indexed)
sheet[2, 2].Value(1234.56m).NumberFormat("$#,##0.00");

sheet.Column("B").Width(18);
sheet.FreezeRows(1);

await wb.SaveAsync("sales.xlsx");
```

Integer cell coordinates are **1-indexed** to match Excel's UI (`sheet[1, 1]` is `A1`).
