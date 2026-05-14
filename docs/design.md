# NetXlsx — Design Document

**Status:** Draft for review
**Date:** 2026-05-14
**Audience:** Public (OSS)

## 1. Mission

Provide an idiomatic, modern, ergonomic C# API for creating and reading `.xlsx` spreadsheets, backed by [NPOI](https://github.com/nissl-lab/npoi). Optimize for the 80% case (tabular data with reasonable styling). Expose NPOI as an escape hatch for the 20% case (charts, pivots, exotic features) without forcing users into it.

## 2. Non-goals

- Replacing NPOI. We wrap it.
- Supporting `.xls`, `.doc`, `.ppt`, or any non-spreadsheet format.
- Re-implementing the OOXML spec.
- Formula *evaluation* (we write formulas; Excel computes them).
- Matching every feature in EPPlus or ClosedXML. We pick what's valuable to typical consumers.

## 3. Foundational decisions

| #  | Decision                          | Choice                                                                       | Why                                                                  |
|----|-----------------------------------|------------------------------------------------------------------------------|----------------------------------------------------------------------|
| 1  | Wrapping philosophy               | Wrapper types + `.Underlying` escape hatch                                   | Clean surface, no feature ceiling, incremental adoption              |
| 2  | API style                         | Fluent-mutable                                                               | Matches domain; reads well; one consistent style                     |
| 3  | Cell addressing                   | `sheet["A1"]` and `sheet[row, col]` both first-class; **1-indexed** for `[r,c]` | Matches Excel UI; eliminates off-by-one when comparing to A1 refs |
| 4  | Style management                  | Auto-deduplication via internal style cache                                  | Biggest ergonomic win over raw NPOI; prevents 64K style budget bug   |
| 5  | Type coercion                     | `Value(object?)` convenience + typed setters (`SetString`, `SetNumber`, …)   | Convenience by default, precision when needed                        |
| 6  | Format scope                      | `.xlsx` only                                                                 | Avoids inheriting all `.xls` constraints; `.xls` users keep raw NPOI |
| 7  | Streaming                         | Separate entry point: `Workbook.CreateStreaming()`                           | Honest about tradeoff (no random access)                             |
| 8  | Reading                           | In scope for v1, including typed row mapping                                 | Write-only is half a library                                         |
| 9  | Async I/O                         | `SaveAsync(stream, ct)`, `OpenAsync` — async-over-sync, documented as such   | Table stakes; revisit when NPOI gains native async                   |
| 10 | Nullability                       | `<Nullable>enable</Nullable>` from commit 1                                  | Modern C#                                                            |
| 11 | Threading                         | Not thread-safe; documented                                                  | Matches NPOI; locking is a footgun                                   |
| 12 | Target frameworks                 | `net8.0; net9.0` (locked)                                                    | Audited 2026-05-14: no .NET Framework / netstandard2.0 consumer demand |
| 13 | AOT / trimming                    | Source generator for typed row mapping; rest is AOT-safe                     | Locks in modern .NET ergonomics                                      |
| 14 | Culture                           | Invariant for serialization; `DisplayCulture` option for format strings only | Spec-correct; prevents bug class                                     |
| 15 | Date system                       | 1900 default; 1904 supported via option; `DateTime.Kind` documented          | Matches Excel default                                                |
| 16 | Stream ownership                  | Leave-open by default; `leaveOpen: false` opt-in                             | BCL convention                                                       |
| 17 | Error model                       | Typed exception hierarchy under `WorkbookException`                          | Idiomatic, granular `catch`                                          |
| 18 | Logging                           | Optional `ILogger` via `WorkbookOptions.Logger`; no hard dep                 | Clean dependency footprint                                           |
| 19 | Telemetry                         | None                                                                         | User trust                                                           |
| 20 | Security                          | Bounded resource options (max uncompressed bytes, sheets, rows)              | Zip bomb defense                                                     |
| 21 | Validation                        | Strict; sanitization helpers available                                       | Fail loud, fix early                                                 |
| 22 | API stability                     | `[Experimental("NETXLSX_NNNN")]` for in-flux APIs                         | Ship and learn                                                       |
| 23 | Versioning                        | Strict SemVer; public-API snapshot tests gate PRs                            | Prevent accidental breaks                                            |
| 24 | NPOI is public dependency         | Yes — `.Underlying` exposes NPOI types                                       | Required for escape hatch; major NPOI bumps = major bumps for us     |
| 25 | Distribution                      | Single NuGet package for v1                                                  | Split later if needed; avoid premature modularity                    |
| 26 | Project name / namespace          | `NetXlsx`                                                                 | Public OSS, NetXlsx-namespaced                                        |
| 27 | Repo location                     | Separate repo at `~/dev/projects/NetXlsx/`                               | Decoupled from NPOI fork; consumes NPOI via NuGet                    |
| 28 | Range enumeration order           | Row-major, populated cells only                                              | Consistent with `ISheet.Rows()`; least surprising for LINQ users     |
| 29 | Color type                        | Owned `NetXlsx.Color` immutable struct; no `System.Drawing` dependency    | `System.Drawing.Common` is Windows-only since .NET 6                 |
| 30 | Async pattern                     | `Task.Run` for `Save`/`Open` (CPU+I/O mixed); no fake-async on trivial sync  | Honest about thread-pool cost; ASP.NET / Blazor users can plan       |
| 31 | Read-side culture                 | Numeric reads parse from invariant; display formatting via `DisplayCulture`  | Excel stores numbers as invariant doubles regardless of display      |
| 32 | Escape hatch concrete types       | `Underlying` returns `XSSFWorkbook`/`XSSFSheet`/`XSSFRow`/`XSSFCell`         | We're `.xlsx`-only; concrete types eliminate caller downcasts        |
| 33 | v2 streaming read                 | Separate `IStreamingSheetReader`; not retrofit on `ISheet`                   | Avoids breaking the materialized `ISheet.Sheets` contract            |

## 4. Performance targets (v1)

| Scenario                                  | Target              |
|-------------------------------------------|---------------------|
| Write 100k rows × 20 cols (in-memory)     | < 3s, < 500 MB      |
| Write 1M rows × 20 cols (streaming)       | < 30s, < 200 MB     |
| Open + read 100k × 20 sheet               | < 4s                |
| Cold create empty workbook + save         | < 50ms              |
| Style dedup overhead — typical (≤ 100 distinct styles) | < 10% vs raw NPOI |
| Style dedup overhead — worst case (high cardinality)   | < 30% vs raw NPOI |

Benchmark suite under `benchmarks/` compares against NPOI direct, EPPlus, ClosedXML.

> **Note on style-dedup target.** The < 10% / < 30% split is provisional. A feasibility benchmark must run before v1.0 lock to validate or revise these numbers.

## 5. Quality gates

- `<Nullable>enable</Nullable>` everywhere; zero warnings.
- 100% XML doc coverage on public API; CI fails on missing.
- Public API snapshot tests via `Microsoft.CodeAnalysis.PublicApiAnalyzers`.
- Golden-file tests for representative workbooks.
- Round-trip tests (write → read → assert).
- Benchmarks run on every PR; regression > 10% fails CI.
- Manual smoke test on Excel (Windows + Mac), LibreOffice, Google Sheets per release.
- Source Link, deterministic builds, symbol packages, signed assemblies.

## 6. v1 interface sketch

Namespace: `NetXlsx`

### 6.1 Entry points

```csharp
public static class Workbook
{
    public static IWorkbook Create(WorkbookOptions? options = null);
    public static IWorkbook CreateStreaming(StreamingOptions? options = null);

    public static IWorkbook Open(string path, WorkbookOptions? options = null);
    public static IWorkbook Open(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null);

    public static Task<IWorkbook> OpenAsync(string path, WorkbookOptions? options = null, CancellationToken ct = default);
    public static Task<IWorkbook> OpenAsync(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null, CancellationToken ct = default);

    public static string SanitizeSheetName(string proposed);
    public static bool IsValidSheetName(string proposed);
}

public sealed class WorkbookOptions
{
    public CultureInfo DisplayCulture { get; init; } = CultureInfo.InvariantCulture;
    public DateSystem DateSystem { get; init; } = DateSystem.Excel1900;
    public ILogger? Logger { get; init; }
    public long MaxUncompressedBytes { get; init; } = 256L * 1024 * 1024;
    public int MaxSheets { get; init; } = 1000;
    public int MaxRowsPerSheet { get; init; } = 1_048_576;
    public string DefaultFontName { get; init; } = "Calibri";
    public double DefaultFontSize { get; init; } = 11;
}

public sealed class StreamingOptions : WorkbookOptions
{
    public int RowAccessWindowSize { get; init; } = 100;
    public bool CompressTempFiles { get; init; } = false;
}

public enum DateSystem { Excel1900, Excel1904 }
```

### 6.2 Workbook

```csharp
public interface IWorkbook : IDisposable, IAsyncDisposable
{
    ISheet AddSheet(string name);
    ISheet AddSheet(string name, int index);
    ISheet this[string name] { get; }
    ISheet this[int index] { get; }
    bool TryGetSheet(string name, [MaybeNullWhen(false)] out ISheet sheet);
    void RemoveSheet(string name);
    void RenameSheet(string oldName, string newName);
    void MoveSheet(string name, int newIndex);

    IReadOnlyList<ISheet> Sheets { get; }

    void Save(string path);
    void Save(Stream stream, bool leaveOpen = true);
    Task SaveAsync(string path, CancellationToken ct = default);
    Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default);

    INamedRange AddNamedRange(string name, string formula);
    IReadOnlyList<INamedRange> NamedRanges { get; }

    NPOI.XSSF.UserModel.XSSFWorkbook Underlying { get; }
}

public interface INamedRange
{
    string Name { get; }
    string Formula { get; }
    string? SheetScope { get; }
}
```

### 6.3 Sheet

```csharp
public interface ISheet
{
    string Name { get; set; }
    int Index { get; }
    IWorkbook Workbook { get; }

    // Cell access (all integer indices are 1-based — see decision #3)
    ICell this[string a1] { get; }                // sheet["A1"]
    ICell this[int row, int col] { get; }         // sheet[1, 1] == sheet["A1"]
    ICell Cell(string a1);
    ICell Cell(int row, int col);                 // 1-based

    // Range access
    IRange Range(string a1Range);                 // sheet.Range("A1:C10")
    IRange Range(int r1, int c1, int r2, int c2); // 1-based, inclusive

    // Rows / columns (1-based)
    IRow Row(int index);                          // sheet.Row(1) == first row
    IColumn Column(int index);                    // sheet.Column(1) == "A"
    IColumn Column(string letter);                // sheet.Column("B")

    IEnumerable<IRow> Rows();                     // populated rows only
    IEnumerable<IRow> Rows(int startInclusive, int endExclusive);

    // Typed row I/O
    void AddRow<T>(T record);                     // appends; uses source-gen mapper if [Worksheet]
    void AddRows<T>(IEnumerable<T> records);
    IEnumerable<T> ReadRows<T>(int headerRow = 0);

    // Layout
    void FreezeRows(int rows);
    void FreezeColumns(int cols);
    void FreezePane(int rows, int cols);

    void MergeCells(string a1Range);
    void UnmergeCells(string a1Range);
    IReadOnlyList<string> MergedRanges { get; }

    void SetTabColor(Color color);
    bool Hidden { get; set; }
    bool ShowGridlines { get; set; }

    int FirstRow { get; }
    int LastRow { get; }

    NPOI.XSSF.UserModel.XSSFSheet Underlying { get; }
}
```

### 6.4 Cell (fluent)

```csharp
public interface ICell
{
    string Address { get; }                       // "A1"
    int RowIndex { get; }                         // 1-based
    int ColumnIndex { get; }                      // 1-based ("A" == 1)

    // Setters (return self for chaining)
    ICell Value(object? value);
    ICell SetString(string value);
    ICell SetNumber(double value);
    ICell SetNumber(decimal value);
    ICell SetDate(DateTime value);
    ICell SetDate(DateOnly value);
    ICell SetBool(bool value);
    ICell SetFormula(string formula);
    ICell Clear();

    // Typed reads
    string? GetString();
    double? GetNumber();
    DateTime? GetDate();
    bool? GetBool();
    string? GetFormula();
    T? GetValue<T>();
    CellKind Kind { get; }

    // Styling (fluent, all return self; mutations go through style cache)
    ICell Bold(bool on = true);
    ICell Italic(bool on = true);
    ICell Underline(UnderlineStyle style = UnderlineStyle.Single);
    ICell Font(string name);
    ICell FontSize(double points);
    ICell FontColor(Color color);
    ICell FontColor(string hex);
    ICell Background(Color color);
    ICell Background(string hex);
    ICell Border(BorderSide sides, BorderStyle style = BorderStyle.Thin, Color? color = null);
    ICell HorizontalAlign(HAlign align);
    ICell VerticalAlign(VAlign align);
    ICell WrapText(bool on = true);
    ICell NumberFormat(string format);            // "$#,##0.00"
    ICell Style(CellStyle style);                 // apply prebuilt
    CellStyle GetStyle();                         // snapshot

    // Annotations
    ICell Comment(string text, string? author = null);
    ICell Hyperlink(string target, string? display = null);

    NPOI.XSSF.UserModel.XSSFCell Underlying { get; }
}

public enum CellKind { Empty, String, Number, Date, Bool, Formula, Error }
[Flags] public enum BorderSide { None=0, Top=1, Right=2, Bottom=4, Left=8, All=15 }
public enum BorderStyle { None, Thin, Medium, Thick, Double, Dashed, Dotted }
public enum HAlign { General, Left, Center, Right, Fill, Justify }
public enum VAlign { Top, Center, Bottom }
public enum UnderlineStyle { None, Single, Double, SingleAccounting, DoubleAccounting }
```

### 6.5 Row / Column / Range

```csharp
public interface IRow : IEnumerable<ICell>
{
    int Index { get; }                            // 1-based
    ISheet Sheet { get; }
    int FirstCol { get; }                         // 1-based
    int LastCol { get; }                          // 1-based
    bool Hidden { get; set; }

    // Dimensions — fluent + property (kept in sync; pick whichever style)
    double HeightPoints { get; set; }
    IRow Height(double points);

    ICell this[int col] { get; }                  // 1-based
    ICell this[string column] { get; }            // row["B"]
    IRow Style(Action<ICell> apply);              // applies to populated cells

    NPOI.XSSF.UserModel.XSSFRow Underlying { get; }
}

public interface IColumn
{
    int Index { get; }                            // 1-based ("A" == 1)
    string Letter { get; }
    ISheet Sheet { get; }
    bool Hidden { get; set; }

    // Dimensions — fluent + property
    double WidthUnits { get; set; }               // Excel column-width units
    IColumn Width(double units);

    IColumn AutoSize();
    IColumn Style(Action<ICell> apply);
    IColumn SetDefaultStyle(CellStyle style);
}

/// <summary>
/// A rectangular range of cells. Enumeration is row-major and yields only
/// cells that are currently populated. To iterate every coordinate in the
/// rectangle regardless of population, use <see cref="EnumerateAll"/>.
/// </summary>
public interface IRange : IEnumerable<ICell>
{
    string Address { get; }
    int FirstRow { get; } int LastRow { get; }    // 1-based, inclusive
    int FirstCol { get; } int LastCol { get; }    // 1-based, inclusive
    int Count { get; }                            // populated cells in range

    IEnumerable<ICell> EnumerateAll();            // dense; materializes empties

    IRange Value(object? value);                  // sets all (dense)
    IRange Style(Action<ICell> apply);            // applies to each populated cell
    IRange Apply(CellStyle style);                // applies to every cell in range (dense)
    IRange Merge();
    IRange ClearContents();
}
```

### 6.6 Color

```csharp
/// <summary>
/// Immutable color value. No dependency on System.Drawing.
/// </summary>
public readonly record struct Color(byte A, byte R, byte G, byte B)
{
    public static Color FromArgb(byte a, byte r, byte g, byte b) => new(a, r, g, b);
    public static Color FromRgb(byte r, byte g, byte b)          => new(0xFF, r, g, b);
    public static Color FromHex(string hex);                     // "#RRGGBB" or "#AARRGGBB"
    public string ToHex();                                       // "#AARRGGBB"

    // Common presets
    public static Color Black  => FromRgb(0, 0, 0);
    public static Color White  => FromRgb(255, 255, 255);
    public static Color Red    => FromRgb(255, 0, 0);
    public static Color Green  => FromRgb(0, 128, 0);
    public static Color Blue   => FromRgb(0, 0, 255);
    public static Color Yellow => FromRgb(255, 255, 0);
    // ... small curated set; full theme support in v2
}
```

### 6.7 Style value object

```csharp
public sealed record CellStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public UnderlineStyle? Underline { get; init; }
    public string? FontName { get; init; }
    public double? FontSize { get; init; }
    public Color? FontColor { get; init; }
    public Color? Background { get; init; }
    public string? NumberFormat { get; init; }
    public HAlign? HAlign { get; init; }
    public VAlign? VAlign { get; init; }
    public bool? WrapText { get; init; }
    public CellBorders? Borders { get; init; }
}

public sealed record CellBorders(
    BorderStyle? Top = null, Color? TopColor = null,
    BorderStyle? Right = null, Color? RightColor = null,
    BorderStyle? Bottom = null, Color? BottomColor = null,
    BorderStyle? Left = null, Color? LeftColor = null);
```

`CellStyle` is a value record. Equal styles produce the same NPOI `ICellStyle` via the workbook's internal style cache.

### 6.8 Typed mapping (source-generated)

```csharp
[Worksheet]
public partial record SalesRecord
{
    [Column("Region", Order = 0)] public required string Region { get; init; }
    [Column("Revenue", Format = "$#,##0.00")] public required decimal Revenue { get; init; }
    [Column("Date", Format = "yyyy-mm-dd")] public required DateOnly Date { get; init; }
    [Ignore] public string InternalNotes { get; init; } = "";
}
```

Source generator emits an `IRowMapper<SalesRecord>` implementation; no reflection at runtime; AOT-safe.

### 6.9 Exception hierarchy

```csharp
public class WorkbookException : Exception { ... }
public sealed class InvalidCellAddressException : WorkbookException { ... }
public sealed class SheetNameException : WorkbookException { ... }
public sealed class StyleBudgetExceededException : WorkbookException { ... }
public sealed class MalformedFileException : WorkbookException { ... }
public sealed class ResourceLimitExceededException : WorkbookException { ... }
public sealed class FormulaException : WorkbookException { ... }
```

## 7. Behavioral notes

### 7.1 Async semantics

NPOI is synchronous. `SaveAsync`/`OpenAsync` use `Task.Run` to offload mixed CPU+I/O work to the thread pool. We do *not* return `Task.FromResult` over trivially-synchronous work; if a method has no async path, it stays synchronous.

Callers in ASP.NET / Blazor contexts should note: `SaveAsync` consumes a thread-pool thread for the duration of the save. If you serialize many workbooks concurrently, throttle.

### 7.2 Culture handling on reads

Excel stores numbers as IEEE doubles regardless of the spreadsheet's locale; reading via `GetNumber()`/`GetValue<double>()` is therefore culture-invariant by construction.

The `WorkbookOptions.DisplayCulture` option affects only:
- How custom number-format strings are interpreted when *writing*.
- How `GetString()` against a date- or number-formatted cell renders the value (when the caller explicitly wants the displayed string rather than the underlying value).

Reading raw values is always invariant. Reading displayed strings is `DisplayCulture`-aware.

### 7.3 Streaming-read forward compatibility

v1 ships streaming *writes* only (`Workbook.CreateStreaming()`). Streaming *reads* are planned for v2 as a separate API:

```csharp
// v2 — sketch
public static class StreamingWorkbook
{
    public static IStreamingSheetReader OpenForRead(string path, StreamingReadOptions? options = null);
    public static IStreamingSheetReader OpenForRead(Stream stream, bool leaveOpen = true, StreamingReadOptions? options = null);
}

public interface IStreamingSheetReader : IDisposable
{
    string SheetName { get; }
    IEnumerable<IStreamingRow> Rows();            // forward-only, single-pass
}
```

This stays *outside* `IWorkbook` deliberately: the materialized `IWorkbook.Sheets` contract cannot honor lazy enumeration without leaking surprises. Keeping streaming-read on its own type preserves v1's API contract without breaking changes.

## 8. Repository layout

```
NetXlsx/
├─ src/
│  ├─ NetXlsx/                  # main library
│  ├─ NetXlsx.SourceGen/        # source generator for [Worksheet] mapping
│  └─ NetXlsx.Analyzers/        # Roslyn analyzers (v2+)
├─ tests/
│  ├─ NetXlsx.Tests/            # unit tests
│  ├─ NetXlsx.GoldenFiles/      # reference workbooks + round-trip tests
│  └─ NetXlsx.PublicApi/        # public-API snapshot tests
├─ benchmarks/
│  └─ NetXlsx.Benchmarks/       # BenchmarkDotNet vs NPOI/EPPlus/ClosedXML
├─ samples/
│  └─ NetXlsx.Cookbook/         # runnable recipes
├─ docs/
│  ├─ design.md                    # this document
│  └─ roadmap.md                   # feature roadmap + binary scope per release
├─ Directory.Build.props
├─ Directory.Packages.props
├─ NetXlsx.sln
└─ README.md
```

## 9. Definition of "extraordinary" for this project

- Code that's a pleasure to read.
- Abstractions that make complex things (style management, typed mapping) simple.
- Error messages that help the caller recover (`InvalidCellAddressException` says *why* the address was invalid, what was expected, and what they passed).
- Performance within 10% of raw NPOI.
- Documentation that clarifies rather than restates.
- The escape hatch is honest: nothing the facade does is uncovertible to a raw-NPOI equivalent.
