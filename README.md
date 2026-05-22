# NetXlsx

Idiomatic C# facade over [NPOI](https://github.com/nissl-lab/npoi) for creating and reading `.xlsx` spreadsheets.

**Status:** **v1.3.0** release-PR ready on `main` (2026-05-22), pending operator tag. v1.3 closes the named-style OOXML round-trip gap (decision I-67) and lands a partial `FilterCriteria.In(...)` for 1–2-value lists (I-68). **v1.2.0** released 2026-05-22 (RemoveTable, table totals row, workbook password, per-column AutoFilter criteria — decisions I-63 through I-66). **v1.1.0** released 2026-05-22 (the 10-slice "common asks" set + fuzz harness + post-review polish — decisions I-50…I-62). **v1.0.0** released 2026-05-20.

The public surface is exercised by **666 tests per TFM × 2 TFMs = 1,332 total runs per CI build** across unit, golden-file, source-generator, fuzz, and public-API snapshot suites. The [CHANGELOG](CHANGELOG.md) has slice-level granularity all the way back to the initial scaffold. The [roadmap](docs/roadmap.md) lists the v1.2 and beyond backlog.

Targets `net8.0` and `net10.0` (both LTS). MIT-licensed.

## Why this exists

NPOI is the only complete OOXML implementation for .NET, but its API is a Java port — it shows. NetXlsx is a thin, opinionated layer on top that:

- **Adds fluent ergonomics.** `sheet.Range("A1:C1").Value("header").Apply(new CellStyle { Bold = true })` instead of the multi-step NPOI dance.
- **Deduplicates styles automatically.** A single internal pool keyed off `CellStyle` value equality. Avoids NPOI's 64K-style cap that bites every team writing many-colored reports (spike-measured at 60–64K — this is a correctness fix, not just polish).
- **Generates typed export/import at compile time.** `[Worksheet]` on a record gets you `sheet.AddRows<T>(items)` / `sheet.ReadRows<T>()` via source generator. No reflection at runtime, AOT-safe in principle.
- **Doesn't hide NPOI.** Every public type exposes `.Underlying` returning the raw `XSSF*` (or `SXSSF*` for streaming) handle. The facade is *additive over NPOI*, not a sandbox.
- **Splits streaming from random-access at the type level.** `Workbook.CreateStreaming()` returns `IStreamingWorkbook` — not the same type as the random-access one. Random-access members are absent from the streaming surface because they'd lie. (Looking at you, EPPlus.)

### How is this different from ClosedXML / EPPlus / MiniExcel?

| | NetXlsx | ClosedXML | EPPlus | MiniExcel |
|---|---|---|---|---|
| Engine | wraps NPOI | own OOXML impl | own OOXML impl | own OOXML impl |
| License | MIT | MIT | Commercial (since 5.0) | Apache-2.0 |
| `.xls` (legacy) | no (explicit `Never`) | no | no | yes |
| Streaming write | yes (SXSSF) | partial | yes | yes |
| Typed export via source gen | yes | no (reflection) | no (reflection) | yes |
| Style auto-dedup | yes (required for correctness past 60K cells) | yes | yes | n/a |
| Formula *evaluation* | no (explicit `Never`) | yes (limited) | yes | no |
| Escape hatch to raw OOXML | yes (`.Underlying`) | partial | partial | no |

**Pick NetXlsx if** you're already using NPOI and want better ergonomics, you write large styled reports (the dedup pool is real), or you want compile-time-checked typed mapping without runtime reflection. **Pick ClosedXML** if you need formula evaluation or want a non-NPOI engine. **Pick MiniExcel** if you need `.xls` support.

## Requirements & known limitations

> **⚠ Not compatible with `PublishAot=true` or `PublishTrimmed=true`.**
> The engine (NPOI 2.7.x) uses `System.Xml.Serialization` and `System.Reflection.Emit` paths that AOT and trim cannot satisfy — measured by [spike 4](spikes/results/spike-4-aot-trim.md). A build-time guard ships with the package: setting either property emits MSBuild error `NXLS0100` / `NXLS0101`. This is a **ceiling imposed by the engine choice**, not by NetXlsx — the block will lift when (a) NPOI 3.x removes those dependencies (re-checked quarterly per [scheduled spikes](docs/scheduled-spikes.md); next 2026-08-16), (b) we pivot to a different engine ([v2 OOXML planning](docs/v2-ooxml-planning.md) tracks bind-ClosedXML and from-scratch options), or (c) NetXlsx implements OOXML directly. **If your deployment requires AOT today**, NetXlsx is not the right choice — ClosedXML is a reasonable alternative that ships AOT-friendly. NetXlsx will revisit this gate at every NPOI 3.x re-check.

> **Not thread-safe by default; opt-in real lock available.** NPOI is not thread-safe; the default facade detection is opportunistic — a reentry counter (decision #43) that *may* surface concurrent mutations as `InvalidOperationException` but has a race-window gap. **Multi-threaded callers should pass `WorkbookOptions { StrictConcurrencyDetection = true }`** (decision I-59) — that mode takes a real per-workbook `Monitor` lock on every mutating path, serializes safely, and accepts a small throughput cost in exchange for a hard "you cannot silently corrupt this workbook" guarantee. The opportunistic default is the right choice for single-threaded callers (zero lock overhead); under concurrent load it will produce flaky "works in tests, fails under load" reports, so flip the option deliberately.

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

using var wb = Workbook.Create();
var sheet = wb.AddSheet("Sales");
sheet.AddRows(records);                       // generator emits the body

foreach (var row in read["Sales"].ReadRows<SalesRow>())
    Console.WriteLine($"{row.Region}: {row.Revenue:C}");
```

The generator (`NetXlsx.SourceGen`) emits stable diagnostic IDs `NXLS0001`–`NXLS0099` for invalid `[Worksheet]` / `[Column]` usage; build-time guards under `NXLS0100`–`NXLS0199` cover AOT/trim incompatibility.

### Escape hatch

When a wrapped operation does not yet exist, every public type exposes its raw NPOI counterpart:

```csharp
XSSFWorkbook raw    = workbook.Underlying;
XSSFSheet    rawSh  = sheet.Underlying;
XSSFRow      rawR   = row.Underlying;
XSSFCell     rawC   = cell.Underlying;
```

This is by design (#1, #32) — the facade is *additive over NPOI*, not a sandbox around it.

See [`samples/NetXlsx.Cookbook`](samples/NetXlsx.Cookbook) for 13 worked recipes covering every public-surface area; each recipe doubles as a golden-file test.

## Documentation

- [Design](docs/design.md) — 52 foundational + 32 implementation decisions, full interface sketch, performance targets, behavioral specifications, quality gates.
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
samples/     NetXlsx.Cookbook (13 worked recipes)
spikes/      NetXlsx.AotSpike + Spike{1,2,3} harnesses + results/
build/       build.sh / build.ps1 (local + CI entry points)
docs/        design, roadmap, implementation-notes, scheduled-spikes, npoi-workarounds
```

## Build

```sh
build/build.sh         # restore + build + test + pack
build/build.sh test    # tests only
build/build.sh bench   # run benchmarks
build/build.sh -- spike-1   # run a specific pre-impl spike
```

PowerShell equivalent: `build/build.ps1`. Both scripts auto-detect a user-level .NET install under `~/.dotnet` and prefer it over a system install (useful if your system SDK is older than 10.x or if you maintain multiple SDK versions side-by-side).

**SDK requirement:** building from source needs the **.NET 10 SDK** (`global.json` pins it with `rollForward: latestFeature`). The .NET 8 + 9 runtimes are needed for the test matrix. See [CONTRIBUTING.md](CONTRIBUTING.md) for install pointers.

## Contributing

Issues and pull requests welcome. The project follows a deliberate design-then-implement loop: substantive new API surface should be discussed against [`docs/design.md`](docs/design.md) before code lands. The [public-API analyzer](src/NetXlsx/PublicAPI.Shipped.txt) gates additions at compile time.

## License

MIT — see [LICENSE](LICENSE).
