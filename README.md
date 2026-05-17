# NetXlsx

Idiomatic C# facade over [NPOI](https://github.com/nissl-lab/npoi) for creating and reading `.xlsx` spreadsheets.

**Status:** pre-1.0, tracking toward v1.0. The main public surface is implemented and exercised by 388 tests (per TFM) across unit, golden-file, and public-API snapshot suites. The [CHANGELOG](CHANGELOG.md) has slice-level granularity; the [roadmap](docs/roadmap.md) lists the remaining v1.0 ship-blockers (streaming write via SXSSF, `WorkbookOptions` for configurable limits, the remaining cookbook recipes, the headless-Linux AutoSize CI job, and the benchmark-regression gate).

Targets `net8.0` and `net9.0`.

## Requirements & known limitations

> **⚠ Not compatible with `PublishAot=true` or `PublishTrimmed=true`.**
> The engine (NPOI 2.7.x) uses `System.Xml.Serialization` and `System.Reflection.Emit` paths that AOT and trim cannot satisfy — measured by [spike 4](spikes/results/spike-4-aot-trim.md). A build-time guard ships with the package: setting either property emits MSBuild error `NXLS0100` / `NXLS0101`. The block will lift when NPOI removes those dependencies (track NPOI 3.x).

> **Not thread-safe.** NPOI is not thread-safe; this facade does not lock. Concurrent mutation produces undefined behavior. A best-effort reentry-counter detection per design decision #43 is in place — concurrent `AddSheet`/mutator calls *may* surface as `InvalidOperationException`. The detection is opportunistic, not a lock; do not rely on it to make concurrent use safe.

> **`IColumn.AutoSize()` requires font metrics.** On headless Linux without `libgdiplus` + a fallback font (e.g. DejaVu), `AutoSize()` throws `MissingFontException` with install commands (design decision I3). The deterministic alternative is `IColumn.Width(double)` with an explicit width.

## What works today

### Workbook lifecycle and round-trip

```csharp
using NetXlsx;

using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");

sheet["A1"].SetString("Region");
sheet[1, 2].SetString("Revenue");           // (row, col), 1-based — A1 == [1,1]

sheet.Row(2).Set(1, "North").Set(2, 1234.56m);

await wb.SaveAsync("sales.xlsx");

using var read = await Workbook.OpenAsync("sales.xlsx");
var name = read["Sales"]["A2"].GetString();    // "North"
var rev  = read["Sales"]["B2"].GetNumber();    // 1234.56
```

### Rows, ranges, columns

```csharp
// Row-level fluent setters (one Set overload per supported scalar type).
sheet.AppendRow().Set(1, "Total").Set(2, 9999.99m);

// Rectangular range — sparse default enumeration; dense via EnumerateAll().
sheet.Range("A1:C1").Value("header").Apply(new CellStyle { Bold = true });
sheet.Range(2, 1, 10, 3).Value(0);            // fill 9x3 with zeros

// Column-level operations.
sheet.Column("B").Width(20);
sheet.Column("C").SetDefaultStyle(new CellStyle { NumberFormat = NumberFormats.Currency });
sheet.Column("D").Hidden = true;
sheet.Column("E").AutoSize();                  // throws MissingFontException on headless-no-fonts
```

### Styling

```csharp
sheet["A1"].Style(new CellStyle
{
    Bold = true,
    FontColor = Color.White,
    Background = Color.FromHex("#003366"),
    HorizontalAlignment = HAlign.Center,
    Borders = CellBorders.All(BorderStyle.Thin, Color.Black),
});

sheet["B2"].NumberFormat(NumberFormats.Currency);
```

`ICell.Style(...)` is a **merge** — non-null axes of the overlay override the existing style; null axes are left untouched. Equal merged styles share one underlying NPOI style index via an internal pool (decision #4), keeping the file under Excel's 64K-style cap even when many cells differ only by background color.

### Dates, times, durations, cell errors

```csharp
sheet["A1"].SetDate(new DateOnly(2026, 5, 16));
sheet["B1"].SetDate(DateTime.Now);
sheet["C1"].SetTime(new TimeOnly(9, 30));
sheet["D1"].SetDuration(TimeSpan.FromMinutes(125));   // [h]:mm:ss format

if (sheet["E1"].Kind == CellKind.Error)
    var err = sheet["E1"].GetError();                  // CellError.DivByZero, .NA, etc.
```

### Freeze panes, merges, hidden sheets

```csharp
sheet.FreezeRows(1);
sheet.MergeCells("A1:C1");                    // anchor value preserved; overlaps throw

var hiddenSheet = wb.AddSheet("Hidden");
hiddenSheet.Hidden = true;
sheet.ShowGridlines = false;
```

### Typed export / import (source-generated)

```csharp
[Worksheet]
public partial record SalesRow(
    [property: Column("Region")]  string Region,
    [property: Column("Revenue", Format = NumberFormats.Currency)] decimal Revenue);

using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");
sheet.AddRows(records);                       // generator emits the body

foreach (var row in read["Sales"].ReadRows<SalesRow>())
    Console.WriteLine($"{row.Region}: {row.Revenue:C}");
```

The generator (`NetXlsx.SourceGen`) emits stable diagnostic IDs `NXLS0001`–`NXLS0006` for invalid `[Worksheet]` / `[Column]` usage.

### Escape hatch

When a wrapped operation does not yet exist, every public type exposes its raw NPOI counterpart:

```csharp
XSSFWorkbook raw    = workbook.Underlying;
XSSFSheet    rawSh  = sheet.Underlying;
XSSFRow      rawR   = row.Underlying;
XSSFCell     rawC   = cell.Underlying;
```

This is by design (#1, #32) — the facade is *additive over NPOI*, not a sandbox around it.

See [`samples/NetXlsx.Cookbook`](samples/NetXlsx.Cookbook) for seven worked recipes that double as golden-file tests.

## Documentation

- [Design](docs/design.md) — 52 foundational + 22 implementation decisions, full v1.0 interface sketch, performance targets, behavioral specifications, quality gates.
- [Roadmap](docs/roadmap.md) — binary feature matrix v1.0 / v1.1 / v2.0 / v3.0 / Never, per-release DoD, process rules.
- [Implementation notes](docs/implementation-notes.md) — patterns and lessons from the implementation phase (not yet a methodology — see file header).
- [Scheduled spikes](docs/scheduled-spikes.md) — quarterly re-checks (e.g. NPOI AOT/trim posture, next due 2026-08-16).
- [NPOI workarounds](docs/npoi-workarounds.md) — catalog of NPOI quirks the facade routes around (currently empty by design).
- [Pre-impl spikes](spikes/results/) — measured outcomes for style-dedup feasibility, streaming back-pressure, async wrapping cost, and AOT/trim posture.
- [Changelog](CHANGELOG.md) — Keep-a-Changelog format, slice-level entries with public-API deltas.

## Layout

```
src/         NetXlsx, NetXlsx.SourceGen
tests/       NetXlsx.Tests, NetXlsx.GoldenFiles, NetXlsx.PublicApi
benchmarks/  NetXlsx.Benchmarks (BenchmarkDotNet)
samples/     NetXlsx.Cookbook
spikes/      NetXlsx.AotSpike + Spike{1,2,3} harnesses + results/
build/       build.sh / build.ps1 (local + CI entry points)
.teamcity/   pipeline-as-code (Kotlin DSL placeholder)
docs/        design, roadmap, implementation-notes, scheduled-spikes, npoi-workarounds
```

## Build

```sh
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
build/build.sh -- spike-1   # run a specific pre-impl spike
```

PowerShell equivalent: `build/build.ps1`. Both scripts auto-detect a user-level .NET install under `~/.dotnet` and prefer it over a system install (lets `net9.0` work on a machine whose system SDK is older).

## Scaffold placeholders (filled before first publish)

- `nuget.config` — public NuGet feed URL.
- `Directory.Build.props` `RepositoryUrl` — github.com/jkindrix.
- `CODEOWNERS` — owning team identifiers.
- `.teamcity/` Kotlin DSL — bound to a real Git host project.
- Source Link package wired (depends on Git host choice).
