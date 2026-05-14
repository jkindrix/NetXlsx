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
| 3  | Cell addressing                   | `sheet["A1"]` and `sheet[row, col]` both first-class; 0-indexed for `[r,c]`  | Cover both user mental models                                        |
| 4  | Style management                  | Auto-deduplication via internal style cache                                  | Biggest ergonomic win over raw NPOI; prevents 64K style budget bug   |
| 5  | Type coercion                     | `Value(object?)` convenience + typed setters (`SetString`, `SetNumber`, …)   | Convenience by default, precision when needed                        |
| 6  | Format scope                      | `.xlsx` only                                                                 | Avoids inheriting all `.xls` constraints; `.xls` users keep raw NPOI |
| 7  | Streaming                         | Separate entry point: `Workbook.CreateStreaming()`                           | Honest about tradeoff (no random access)                             |
| 8  | Reading                           | In scope for v1, including typed row mapping                                 | Write-only is half a library                                         |
| 9  | Async I/O                         | `SaveAsync(stream, ct)`, `OpenAsync` — async-over-sync, documented as such   | Table stakes; revisit when NPOI gains native async                   |
| 10 | Nullability                       | `<Nullable>enable</Nullable>` from commit 1                                  | Modern C#                                                            |
| 11 | Threading                         | Not thread-safe; documented                                                  | Matches NPOI; locking is a footgun                                   |
| 12 | Target frameworks                 | `net8.0; net9.0`                                                             | Modern .NET; revisit `netstandard2.0` only on request                |
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

## 4. Performance targets (v1)

| Scenario                                  | Target              |
|-------------------------------------------|---------------------|
| Write 100k rows × 20 cols (in-memory)     | < 3s, < 500 MB      |
| Write 1M rows × 20 cols (streaming)       | < 30s, < 200 MB     |
| Open + read 100k × 20 sheet               | < 4s                |
| Cold create empty workbook + save         | < 50ms              |
| Style dedup overhead (1M cell writes)     | < 10% vs raw NPOI   |

Benchmark suite under `benchmarks/` compares against NPOI direct, EPPlus, ClosedXML.

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
    bool TryGetSheet(string name, out ISheet sheet);
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

    NPOI.SS.UserModel.IWorkbook Underlying { get; }
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

    // Cell access
    ICell this[string a1] { get; }                // sheet["A1"]
    ICell this[int row, int col] { get; }         // sheet[0, 0]
    ICell Cell(string a1);
    ICell Cell(int row, int col);

    // Range access
    IRange Range(string a1Range);                 // sheet.Range("A1:C10")
    IRange Range(int r1, int c1, int r2, int c2);

    // Rows / columns
    IRow Row(int index);
    IColumn Column(int index);
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

    NPOI.SS.UserModel.ISheet Underlying { get; }
}
```

### 6.4 Cell (fluent)

```csharp
public interface ICell
{
    string Address { get; }                       // "A1"
    int RowIndex { get; }
    int ColumnIndex { get; }

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

    NPOI.SS.UserModel.ICell Underlying { get; }
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
    int Index { get; }
    ISheet Sheet { get; }
    int FirstCol { get; }
    int LastCol { get; }
    double Height { get; set; }
    bool Hidden { get; set; }
    ICell this[int col] { get; }
    ICell this[string column] { get; }            // row["B"]
    IRow Style(Action<ICell> apply);              // applies to populated cells
    NPOI.SS.UserModel.IRow Underlying { get; }
}

public interface IColumn
{
    int Index { get; }
    string Letter { get; }
    ISheet Sheet { get; }
    double Width { get; set; }
    bool Hidden { get; set; }
    void AutoSize();
    IColumn Style(Action<ICell> apply);
    void SetDefaultStyle(CellStyle style);
}

public interface IRange : IEnumerable<ICell>
{
    string Address { get; }
    int FirstRow { get; } int LastRow { get; }
    int FirstCol { get; } int LastCol { get; }
    int Count { get; }

    IRange Value(object? value);                  // sets all
    IRange Style(Action<ICell> apply);            // applies to each
    IRange Apply(CellStyle style);
    IRange Merge();
    IRange ClearContents();
}
```

### 6.6 Style value object

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

### 6.7 Typed mapping (source-generated)

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

### 6.8 Exception hierarchy

```csharp
public class WorkbookException : Exception { ... }
public sealed class InvalidCellAddressException : WorkbookException { ... }
public sealed class SheetNameException : WorkbookException { ... }
public sealed class StyleBudgetExceededException : WorkbookException { ... }
public sealed class MalformedFileException : WorkbookException { ... }
public sealed class ResourceLimitExceededException : WorkbookException { ... }
public sealed class FormulaException : WorkbookException { ... }
```

## 7. Repository layout

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

## 8. Definition of "extraordinary" for this project

- Code that's a pleasure to read.
- Abstractions that make complex things (style management, typed mapping) simple.
- Error messages that help the caller recover (`InvalidCellAddressException` says *why* the address was invalid, what was expected, and what they passed).
- Performance within 10% of raw NPOI.
- Documentation that clarifies rather than restates.
- The escape hatch is honest: nothing the facade does is uncovertible to a raw-NPOI equivalent.
