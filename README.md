# NetXlsx

Idiomatic C# library for creating and reading `.xlsx` spreadsheets, built on Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK).

**Status:** **v2.0.1** — released 2026-06-04. The v2.0.0 engine swap (decision I-82) shipped: the engine is Microsoft's Open XML SDK, NPOI is removed from the library, and the escape hatches are SDK-typed (a breaking change for `.Underlying` consumers; see the [CHANGELOG](CHANGELOG.md) migration table). v2.0.1 follows immediately with the bulk-write performance fix (I-87) and is the first publishable v2 version. **v1.3.0**, **v1.2.0** and **v1.1.0** released 2026-05-22; **v1.0.0** released 2026-05-20.

The public surface is exercised on every CI build, across both target frameworks, by five suites — unit, golden-file, source-generator, fuzz, and public-API snapshot. (An exact count isn't quoted here: it's not gated, so a hand-maintained number only drifts. Run `bash build/build.sh test` for the current tally.) The [CHANGELOG](CHANGELOG.md) has slice-level granularity all the way back to the initial scaffold. The [roadmap](docs/roadmap.md) lists the v1.2 and beyond backlog.

Targets `net8.0` and `net10.0` (both LTS). MIT-licensed.

## Why this exists

Raw OOXML is verbose and easy to get subtly wrong; the Open XML SDK is schema-complete but deliberately low-level. NetXlsx is a thin, opinionated layer on top that (v1.x wrapped NPOI; the v2.0.0 engine swap, decision I-82, moved to the SDK with the same public surface):

- **Adds fluent ergonomics.** `sheet.Range("A1:C1").Value("header").Apply(new CellStyle { Bold = true })` instead of hand-building schema elements.
- **Deduplicates styles automatically.** A single internal pool keyed off `CellStyle` value equality. Avoids Excel's 64K-style cap that bites every team writing many-colored reports (spike-measured at 60–64K — this is a correctness fix, not just polish).
- **Generates typed export/import at compile time.** `[Worksheet]` on a record gets you `sheet.AddRows(items)` / `sheet.ReadRows()` extension methods via source generator — per-type, no generic dispatch. No reflection at runtime, AOT-safe in principle.
- **Doesn't hide the OOXML.** Every random-access public type exposes `.Underlying` returning the raw Open XML SDK object (`SpreadsheetDocument`, `Worksheet`, `Cell`, parts for charts/tables). The facade is *additive over the document*, not a sandbox.
- **Splits streaming from random-access at the type level.** `Workbook.CreateStreaming()` returns `IStreamingWorkbook` — not the same type as the random-access one. Random-access members are absent from the streaming surface because they'd lie. (Looking at you, EPPlus.)

### How is this different from ClosedXML / EPPlus / MiniExcel?

| | NetXlsx | ClosedXML | EPPlus | MiniExcel |
|---|---|---|---|---|
| Engine | Open XML SDK (Microsoft) | Open XML SDK (Microsoft) | own OOXML impl | own OOXML impl |
| License | MIT | MIT | Commercial (since 5.0) | Apache-2.0 |
| `.xls` (legacy) | no (explicit `Never`) | no | no | yes |
| Streaming write | yes (forward-only, bounded memory) | partial | yes | yes |
| Typed export via source gen | yes | no (reflection) | no (reflection) | yes |
| Style auto-dedup | yes (required for correctness past 60K cells) | yes | yes | n/a |
| Formula *evaluation* | no (explicit `Never`) | yes (limited) | yes | no |
| Escape hatch to raw OOXML | yes (`.Underlying`) | partial | partial | no |

**Pick NetXlsx if** you want a deliberately thin, AOT-friendly layer (ClosedXML shares the SDK engine but builds a much larger object model on it — the differentiation is thinness, streaming write, source-gen typed mapping, and the raw-SDK escape hatch, not the engine), you write large styled reports (the dedup pool is real), or you want compile-time-checked typed mapping without runtime reflection. **Pick ClosedXML** if you need formula evaluation. **Pick MiniExcel** if you need `.xls` support.

## Requirements & known limitations

> **✅ AOT- and trim-compatible since v2.0.0.** The engine (Microsoft's Open XML SDK) passed the `PublishTrimmed` + `PublishAot` audit at the v2.0.0 cutover: zero IL/AOT warnings, and a representative create/style/save/open workload runs correctly as a trimmed deployment and as a native binary. The pre-v2 build-time block (`NXLS0100` / `NXLS0101`, a ceiling imposed by the retired NPOI engine — measured by [spike 4](spikes/results/spike-4-aot-trim.md)) has been removed. The typed-mapping source generator was always AOT-safe by construction (compile-time emit, no runtime reflection).

> **Actively-maintained engine since v2.0.0.** The pre-v2 "deliberately frozen engine" posture (NPOI pinned at 2.7.3 for license reasons; upstream security patches not flowing in) no longer applies: the engine is `DocumentFormat.OpenXml` — MIT-licensed and maintained by Microsoft — and upstream fixes flow through normal version bumps. See [SECURITY.md](SECURITY.md) for the current dependency posture.

> **Not thread-safe by default; opt-in real lock available.** Workbooks are not thread-safe; the default facade detection is opportunistic — a reentry counter (decision #43) that *may* surface concurrent mutations as `InvalidOperationException` but has a race-window gap. **Multi-threaded callers should pass `WorkbookOptions { StrictConcurrencyDetection = true }`** (decision I-59) — that mode takes a real per-workbook `Monitor` lock on every mutating path, serializes safely, and accepts a small throughput cost in exchange for a hard "you cannot silently corrupt this workbook" guarantee. The opportunistic default is the right choice for single-threaded callers (zero lock overhead); under concurrent load it will produce flaky "works in tests, fails under load" reports, so flip the option deliberately.

> **`IColumn.AutoSize()` is deterministic and headless-safe since v2.0.0** (decision I-84): widths are measured against embedded numeric font-metric tables (metric-compatible SIL-OFL twins of Calibri/Arial/Times/Courier families) — no fontconfig, no `libgdiplus`, identical results on every machine. Fonts outside the embedded set throw `MissingFontException` naming the font; the fallback is `IColumn.Width(double)` with an explicit width.

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

`ICell.Style(...)` is a **merge** — non-null axes of the overlay override the existing style; null axes are left untouched. Equal merged styles share one underlying `cellXfs` index via an internal pool (decision #4), keeping the file under Excel's 64K-style cap even when many cells differ only by background color.

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

### Streaming write (large workbooks)

```csharp
// Use streaming once you're past ~30k rows — spike-2-measured threshold
// where in-memory writes exceed the design's memory budget.
await using var wb = Workbook.CreateStreaming(new StreamingOptions { RowAccessWindowSize = 1_000 });
var sheet = wb.AddSheet("BigData");

for (int r = 1; r <= 1_000_000; r++)
    sheet.AppendRow().Set(1, r).Set(2, $"row-{r}").Set(3, r * 1.5);

await wb.SaveAsync("big.xlsx");
```

`IStreamingWorkbook` is a deliberately narrower contract — random-access members are absent because they'd lie once a row is flushed past the window.

### Formulas and named ranges

```csharp
sheet["A1"].SetNumber(10);
sheet["A2"].SetNumber(20);
sheet["A3"].SetFormula("=SUM(A1:A2)");        // leading '=' optional

wb.AddNamedRange("MonthlySales", "Data!$A$1:$A$12");
sheet["B1"].SetFormula("=SUM(MonthlySales)"); // readable formulas via named ranges
```

Formula *evaluation* is intentionally out of scope — Excel and other competent consumers recalculate on open. NetXlsx never pre-computes cached values (design §7.8).

### Comments and hyperlinks

```csharp
sheet["A1"].Comment("Reviewer flagged for follow-up");          // default author "NetXlsx"
sheet["B2"].Hyperlink("https://example.com", display: "Docs");  // sniffed scheme
```

### Typed export / import (source-generated)

```csharp
[Worksheet]
public partial record SalesRow(
    [property: Column("Region")]  string Region,
    [property: Column("Revenue", Format = NumberFormats.Currency)] decimal Revenue);

using (var wb = Workbook.Create())
{
    var sheet = wb.AddSheet("Sales");
    sheet.AppendRow().Set(1, "Region").Set(2, "Revenue");  // header row (authored by hand)
    sheet.AddRows(records);                                // generator emits the body
    wb.Save("sales.xlsx");
}

using var read = Workbook.Open("sales.xlsx");
foreach (var row in read["Sales"].ReadRows())              // maps columns by [Column] header name
    Console.WriteLine($"{row.Region}: {row.Revenue:C}");
```

The generated methods are per-type extensions (`SalesRow_SheetExtensions.AddRows/ReadRows`), not a generic `ReadRows<T>()` — with several `[Worksheet]` types in scope, disambiguate by calling through the generated static class.

The generator (`NetXlsx.SourceGen`) emits stable diagnostic IDs `NXLS0001`–`NXLS0099` for invalid `[Worksheet]` / `[Column]` usage. (The `NXLS0100`–`NXLS0199` AOT/trim build guards were retired at v2.0.0 — the engine passed the AOT/trim audit.)

### Escape hatch

When a wrapped operation does not yet exist, every random-access public type exposes its raw Open XML SDK counterpart:

```csharp
SpreadsheetDocument doc   = workbook.Underlying;  // the live document + part graph
Worksheet           rawSh = sheet.Underlying;     // worksheet DOM root
Row                 rawR  = row.Underlying;
Cell                rawC  = cell.Underlying;      // materializes the node on access
ChartPart           chart = myChart.Underlying;   // chart/table content lives in its own part
```

This is by design (#1, #32 / I-82) — the facade is *additive over the OOXML document*, not a sandbox around it. (The streaming surface has no hatch: rows stream forward-only and the package exists only at `Save`.)

See [`samples/NetXlsx.Cookbook`](samples/NetXlsx.Cookbook) for worked recipes spanning the v1.0 through v2.0 public surface — from hello-world and typed import/export through conditional formatting, charts, sorting, panes, grouping, shapes/connectors, picture borders, totals rows, AutoFilter criteria, named-style integration, and `.xlsm` passthrough. Each recipe doubles as a golden-file test (an exact count isn't quoted here, for the same no-drift reason as test counts — the directory and `cookbook --help` are the live tally).

## Documentation

- [Design](docs/design.md) — the numbered decision record (52 foundational decisions plus the growing I-NN implementation series; an exact tally isn't quoted here for the same no-drift reason as test counts), full interface sketch, performance targets, behavioral specifications, quality gates.
- [Roadmap](docs/roadmap.md) — binary feature matrix v1.0 / v1.1 / v2.0 / v3.0 / Never, per-release DoD, process rules.
- [Interop & limits](docs/interop.md) — the LibreOffice resave matrix (asserted nightly in CI), size ceilings, formula-injection guidance, metadata posture.
- [Implementation notes](docs/implementation-notes.md) — patterns and lessons from the implementation phase (not yet a methodology — see file header).
- [Scheduled spikes](docs/scheduled-spikes.md) — quarterly re-checks (historical: the NPOI AOT/trim re-check is retired with the v2.0.0 engine swap).
- [NPOI workarounds](docs/npoi-workarounds.md) — historical: the NPOI quirk catalog, retired empty at the v2.0.0 engine swap.
- [Pre-impl spikes](spikes/results/) — measured outcomes for style-dedup feasibility, streaming back-pressure, async wrapping cost, and AOT/trim posture.
- [Changelog](CHANGELOG.md) — Keep-a-Changelog format, slice-level entries with public-API deltas.

## Layout

```
src/         NetXlsx, NetXlsx.SourceGen
tests/       NetXlsx.Tests, NetXlsx.GoldenFiles, NetXlsx.Fuzz, NetXlsx.PublicApi
benchmarks/  NetXlsx.Benchmarks (BenchmarkDotNet)
samples/     NetXlsx.Cookbook (worked recipes — v1.0 through v2.0 surface; each a golden test)
spikes/      NetXlsx.AotSpike + Spike{1,2,3} harnesses + results/
build/       build.sh / build.ps1 (local + CI entry points)
docs/        design, roadmap, implementation-notes, scheduled-spikes,
             remediation ledger, historical NPOI docs
```

## Build

```sh
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
build/build.sh -- spike-1   # run a specific pre-impl spike
```

PowerShell equivalent: `build/build.ps1`. Both scripts auto-detect a user-level .NET install under `~/.dotnet` and prefer it over a system install (useful if your system SDK is older than 10.x or if you maintain multiple SDK versions side-by-side).

**SDK requirement:** building from source needs the **.NET 10 SDK** (`global.json` pins it with `rollForward: latestFeature`). The .NET 8 runtime is needed for the test matrix. See [CONTRIBUTING.md](CONTRIBUTING.md) for install pointers.

## Contributing

Issues and pull requests welcome. The project follows a deliberate design-then-implement loop: substantive new API surface should be discussed against [`docs/design.md`](docs/design.md) before code lands. The [public-API analyzer](src/NetXlsx/PublicAPI.Shipped.txt) gates additions at compile time.

## License

MIT — see [LICENSE](LICENSE).
