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
| 7  | Streaming                         | Separate **type** (`IStreamingWorkbook`) and entry point (`Workbook.CreateStreaming()`) | Random-access members would lie under streaming; type split makes the contract honest |
| 8  | Reading                           | In scope for v1, including typed row mapping                                 | Write-only is half a library                                         |
| 9  | Async I/O                         | `SaveAsync(stream, ct)`, `OpenAsync` — async-over-sync, documented as such   | Table stakes; revisit when NPOI gains native async                   |
| 10 | Nullability                       | `<Nullable>enable</Nullable>` from commit 1                                  | Modern C#                                                            |
| 11 | Threading                         | Not thread-safe; documented                                                  | Matches NPOI; locking is a footgun                                   |
| 12 | Target frameworks                 | `net8.0; net10.0` as of v1.0.0 (net10.0 added 2026-05-20 per I22; net9.0 dropped at the v1.0.0 tag per I24, after its 2026-05-12 EOS) | Audited 2026-05-14: no .NET Framework / netstandard2.0 consumer demand. Both shipped TFMs are LTS. |
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
| 34 | Typed-mapping API surface         | Source-generated **extension methods** on `ISheet`, not methods on `ISheet`  | No `[Worksheet]` ⇒ no extension ⇒ compile-time error; AOT-safe by construction |
| 35 | Open-side style handling          | Existing styles imported by reference; dedup cache only acts on facade-created styles | No-op `Open → Save` produces no styles-part churn (diff-friendly) |
| 36 | `decimal` write semantics         | `SetNumber(decimal)` rounds via `(double)d`; exact textual preservation requires `SetString` | `.xlsx` stores IEEE-754 doubles only; no silent text fallback |
| 37 | Excel hard-limit enforcement      | Throw `ResourceLimitExceededException` on overflow (rows, cols, string length, etc.) | Fail loud at the boundary, never silently truncate |
| 38 | Cancellation under `Task.Run`     | `CancellationToken` honored only at offload boundaries, not mid-NPOI         | NPOI is not cancellation-aware; over-promising would mislead callers |
| 39 | Format-string culture handling    | Format strings are pass-through bytes; rendering is Excel's responsibility   | Spec-correct; clarifies that `DisplayCulture` only affects read-side displayed strings |
| 40 | Cell-materialization semantics    | Accessing an unwritten cell auto-materializes as blank (`CellKind.Empty`)    | Matches Excel's "every cell exists" mental model; mirrors NPOI `CREATE_NULL_AS_BLANK` |
| 41 | Sheet-name collision policy       | `AddSheet` throws `SheetNameException`; `Workbook.SuggestSheetName` available | Auto-suffix is surprising at the library layer; helper makes the intent explicit |
| 42 | Disposal contract                 | Double-dispose: safe no-op. Use-after-dispose: `ObjectDisposedException`. Streaming dispose without prior `Save`: discard, no auto-save | BCL convention; auto-save on dispose silently hides bugs |
| 43 | Concurrent-access policy          | Detect concurrent mutation via reentry counter; throw `InvalidOperationException` | Fail loud rather than corrupt; we explicitly decline to add locks (#11) |
| 44 | Round-trip preservation           | All OPC parts we don't model are preserved verbatim on `Open → Save`         | Pivot caches, custom XML, conditional formatting, vendor extensions must survive untouched; protects user data |
| 45 | String storage policy             | Shared strings by default for in-memory workbooks; inline strings for streaming writes | Matches Excel and NPOI defaults; streaming inline avoids holding the SST in memory |
| 46 | Formula value cache               | Formulas written with no cached value; Excel recalculates on open            | Pre-computing via NPOI's evaluator is slow, fragile, and diverges from Excel for non-trivial formulas |
| 47 | OOXML schema variant              | Read Strict + Transitional; write Transitional                               | Strict is rare in the wild; Transitional is what Excel emits |
| 48 | Time / duration support           | `ICell.SetTime(TimeOnly)` and `SetDuration(TimeSpan)`; stored as Excel time fraction | Common need; cheap to add; format-string responsibility remains the caller's |
| 49 | Cell error values                 | `CellError` enum + `ICell.GetError()`; `CellKind.Error` already present       | Reading error states is a real need; without this users would have to dive into `.Underlying` |
| 50 | Stream position contract on Open  | Stream must be readable and at position 0; we do not `Seek`                  | Explicit position requirements are standard BCL practice (cf. `XmlReader`, `ZipArchive`) |
| 51 | Wrong-format file on Open         | Opening a non-`.xlsx` file (e.g., `.xls`) throws `MalformedFileException` with format hint | Fail loud with an actionable message; never attempt to coerce |
| 52 | Read-side thread safety           | Concurrent reads of a workbook are not safe; reads + any mutation throw       | NPOI is not thread-safe for reads either; consistent with #43 |

### 3.1 Implementation-level decisions

These are below the API contract but above implementation discretion. They are decisions an implementer would otherwise have to make ad-hoc.

| #   | Decision                       | Choice                                                                       | Rationale                                                          |
|-----|--------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------|
| I1  | NPOI version pin               | `2.7.x` — exact patch chosen at scaffold (current candidate `2.7.3`); pinned in `Directory.Packages.props` | One source of truth; bumped via PR with golden-file diff review     |
| I2  | AOT / trim posture             | **AOT-incompatible AND trim-incompatible.** Measured by spike 4 (2026-05-15) against NPOI 2.7.3: both `PublishAot=true` and `PublishTrimmed=true` produce binaries that throw `POIXMLException` at `XSSFWorkbook` init. The facade layer is AOT-clean by construction; the engine is not. Re-evaluate when NPOI removes its `System.Xml.Serialization` / `System.Reflection.Emit` dependencies | Spike-measured. See `spikes/results/spike-4-aot-trim.md` |
| I3  | `AutoSizeColumn` on headless Linux | Ship with documented font dependency; if `libgdiplus` + a fallback font are unavailable, `AutoSizeColumn` throws `MissingFontException : WorkbookException` with installation guidance | NPOI's column-sizing requires font metrics; failing loud is better than silently producing wrong widths |
| I4  | A1 parser — accepted forms     | See §6.11; canonical form returned by `ICell.Address` is uppercase, no `$`, no sheet prefix | Real callers paste many forms; canonicalizing on output makes diffs stable |
| I5  | Source generator architecture  | `IIncrementalGenerator`; cross-assembly `[Worksheet]` types are ignored (only the current compilation is scanned); diagnostic catalog under `NXLS0001+` (see §6.12) | Modern Roslyn; same scoping rule as `System.Text.Json`; explicit diagnostic IDs prevent renumbering churn |
| I6  | `GetValue<T>()` mismatch       | Returns `null` (for `T?`) or `default(T)` (for non-nullable `T`) when conversion fails; never throws on type mismatch. Throws only on dispose / sheet-removed states | Predictable; lets callers chain `?? fallback`; matches `IDataReader.GetValueOrDefault` convention |
| I7  | `GetString()` over non-string  | Returns the **displayed** form: formula → cached result as string; error → error code (`"#DIV/0!"`); empty → `""`; bool → invariant `"TRUE"` / `"FALSE"` | Matches what the user sees in Excel; never throws |
| I8  | `IRange` method semantics      | `Value(object?)` and `Apply(CellStyle)` are **dense** (materialize empties). The sparse iteration is renamed `ForEachPopulated(Action<ICell>)` — see §6.6 | Eliminates the populated-vs-dense footgun across near-identical method names |
| I9  | Named-range scope              | `AddNamedRange(name, formula, sheetScope = null)` — `null` = workbook scope; non-null = sheet-scoped | Single overload; `null` is the common case |
| I10 | Default font / size seam       | `WorkbookOptions.DefaultFontName/Size` populate the workbook's default cell style (style index 0). Any cell without an explicit font inherits transitively. Setting them after sheets exist is a no-op | Matches Excel's model; single seam, well-defined |
| I11 | Comment author default         | `"NetXlsx"` when no author is passed                                       | Explicit attribution; no PII leak via `Environment.UserName`        |
| I12 | Built-in number formats        | A `NumberFormats` static class enumerates the v1.0 set — see §6.13            | Frozen at v1.0; additions in v1.1+ require explicit decision        |
| I13 | Hyperlink target sniffing      | Scheme-sniffed: `http(s)://`, `mailto:`, `file://`, internal `#Sheet!Range` syntax. Anything else throws `ArgumentException` | Simpler than an explicit `HyperlinkKind` parameter; covers the common cases |
| I14 | Stream-on-Open requirements    | Stream must be readable, seekable, and at position 0; otherwise `ArgumentException` (decision #50 set position; this completes the contract) | NPOI's `XSSFWorkbook(Stream)` copies to temp; seekable is required upstream |
| I15 | Negative `TimeSpan`            | `SetDuration(TimeSpan)` throws `ArgumentOutOfRangeException` on negative values | Excel cannot render negative time; silent storage is a worse bug   |
| I16 | Boolean display culture        | `GetString()` on a boolean returns invariant `"TRUE"` / `"FALSE"` regardless of `DisplayCulture` | Booleans are stored as `b="1"/"0"` in OOXML; localized display is a UI concern, not a library concern |
| I17 | `DateTime.Kind` handling       | Stored as-is; no timezone conversion on write. Reads always return `Kind = Unspecified`. Documented in §7.10 | Excel has no timezone concept; converting silently would lose information |
| I18 | Test fixture provenance        | Two categories: (a) hand-crafted in Excel — stored as binary, source documented per-fixture in `tests/NetXlsx.GoldenFiles/README.md`; (b) script-generated — each has a sibling `.gen.cs` that produces it on demand | Reproducibility when NPOI bumps change byte output; binary fixtures stay traceable |
| I19 | Pre-1.0 tag cadence            | `v0.x.y` tags through scaffold and v1.0 implementation; `v1.0.0` only after the v1.0 DoD passes. Internal consumers can take preview deps on `v0.x.y` | Honest progression; preview consumers know the deal |
| I20 | `PublicAPI.Shipped/Unshipped` policy | At every tagged release, the release PR moves the `Unshipped.txt` content into `Shipped.txt` | Single, mechanical, reviewable transition                            |
| I21 | Spike-failure handling         | If a pre-impl spike misses its target, the architect (current: project owner) decides among: revise target, implement workaround, descope. Decision recorded as a new design-doc revision before code lands | No silent target erosion |
| I22 | Future TFM policy              | New TFMs (`net10.0+`) added in the next minor release after the TFM reaches GA; never in a patch release | Predictable consumer impact |
| I23 | NPOI version pin policy        | Pin at **NPOI 2.7.3** (last Apache-2.0 release). Do not bump to 2.8+ while the Open Source Maintenance Fee (OSMF) EULA is in effect on binary releases. Re-evaluate quarterly per `docs/scheduled-spikes.md`. | NPOI 2.8.0 added an OSMF EULA requiring revenue-generating consumers (≥ US $10K/year) to pay a monthly maintenance fee. NetXlsx is MIT-licensed and would not pass that obligation cleanly to downstream consumers without a loud disclosure. Pinning preserves the project's "MIT all the way down" promise; quarterly re-check keeps the option open. The long-term resolution (decision-deferred) is to implement OOXML directly and drop the NPOI dependency entirely — see `docs/long-term.md`. |
| I24 | TFM support window policy      | Support **latest LTS + previous LTS + current STS** while the previous LTS is in support. When a TFM reaches end-of-support, drop it at the **next minor release**. Pre-1.0 keeps STS TFMs through the v1.0 cutting window to ease adoption; the drop happens at the v1.0 tag itself. | Microsoft's STS support window (18 months) is shorter than the typical consumer migration window. Keeping STS TFMs past their EOS imposes test-matrix cost on us without clear benefit; dropping immediately at EOS imposes a re-targeting cost on consumers mid-migration. The minor-release alignment splits the difference and gives consumers a CHANGELOG signal of when their TFM choice changes. v1.0 drop of net9.0 (EOS 2026-05-12) is the first application of this policy. |

## 4. Performance targets (v1)

| Scenario                                  | Target              |
|-------------------------------------------|---------------------|
| Write 30k rows × 20 cols (in-memory)      | < 3s, < 500 MB ΔWS  |
| Write 1M rows × 20 cols (streaming)       | < 30s, < 200 MB ΔWS |
| Open + read 100k × 20 sheet               | < 4s                |
| Cold create empty workbook + save         | < 50ms              |
| Style pool size                           | equals count of distinct `CellStyle` values, regardless of styled-cell count |
| Styled-write throughput (small palette)   | > 500k styled cells/s |
| Styled-write throughput (1k-palette)      | > 400k styled cells/s |

Benchmark suite under `benchmarks/` compares against NPOI direct, EPPlus, ClosedXML.

> **Spike-measured numbers.** The in-memory row-count target (30k) and the style throughput targets above are spike-derived (spike 1 + spike 2, 2026-05-15) — not optimistic guesses. The original "100k rows in-memory" target was missed by ~2×; the threshold was lowered to a value that holds. Callers needing more than ~30k rows should use the streaming entry point (`Workbook.CreateStreaming()`), which sustained a flat ~70 MB ΔGC at 500k rows × 20 cols in spike 2.
>
> **The "raw NPOI without dedup" baseline does not exist.** NPOI hits its ~60,000-style cap on any workbook with > ~60k styled cells when each cell gets a fresh `ICellStyle`. Style dedup is therefore the only viable path; the original "<10% / <30% overhead" framing was measuring a phantom and has been replaced with absolute capacity + throughput targets (above).
>
> **Interim style cache (date/time slice, decision S29).** Until the full style-pool dedup arrives with the styling API slice, `XssfWorkbook` keeps four lazily-allocated `ICellStyle` instances for the date/time-default format strings. These four styles will register as dedup entries once the full pool exists; they don't compete with it. Style allocations from `SetDate`/`SetTime`/`SetDuration` are therefore O(1) per workbook regardless of cell count, even in the interim.

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
    public static IStreamingWorkbook CreateStreaming(StreamingOptions? options = null);

    public static IWorkbook Open(string path, WorkbookOptions? options = null);
    public static IWorkbook Open(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null);

    public static Task<IWorkbook> OpenAsync(string path, WorkbookOptions? options = null, CancellationToken ct = default);
    public static Task<IWorkbook> OpenAsync(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null, CancellationToken ct = default);

    public static string SanitizeSheetName(string proposed);
    public static bool IsValidSheetName(string proposed);

    /// <summary>
    /// Returns <paramref name="proposed"/> if no sheet with that name exists in
    /// <paramref name="wb"/>; otherwise returns "<paramref name="proposed"/> (2)",
    /// "(3)", etc., until an unused name is found. Truncates to the 31-char limit.
    /// </summary>
    public static string SuggestSheetName(IWorkbook wb, string proposed);
}

public sealed class WorkbookOptions
{
    public CultureInfo DisplayCulture { get; init; } = CultureInfo.InvariantCulture;
    public DateSystem DateSystem { get; init; } = DateSystem.Excel1900;
    public ILogger? Logger { get; init; }
    // Read-side safety (zip-bomb defense; not applied on write):
    public long ReadMaxUncompressedBytes { get; init; } = 256L * 1024 * 1024;
    public int ReadMaxSheets { get; init; } = 1000;

    // Excel hard limits (write-side enforcement):
    public int MaxRowsPerSheet { get; init; } = 1_048_576;     // Excel hard cap
    public int MaxColsPerSheet { get; init; } = 16_384;        // Excel hard cap ("XFD")
    public int MaxCellTextLength { get; init; } = 32_767;      // Excel hard cap
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

    INamedRange AddNamedRange(string name, string formula, string? sheetScope = null);
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

### 6.3 Streaming workbook (write-only)

A deliberately narrower contract than `IWorkbook`. Random-access members are absent — once a row is flushed, it cannot be revisited.

```csharp
public interface IStreamingWorkbook : IDisposable, IAsyncDisposable
{
    IStreamingSheet AddSheet(string name);

    void Save(string path);
    void Save(Stream stream, bool leaveOpen = true);
    Task SaveAsync(string path, CancellationToken ct = default);
    Task SaveAsync(Stream stream, bool leaveOpen = true, CancellationToken ct = default);

    NPOI.XSSF.Streaming.SXSSFWorkbook Underlying { get; }
}

public interface IStreamingSheet
{
    string Name { get; }
    IStreamingWorkbook Workbook { get; }
    IStreamingRow AppendRow();                    // creates next row; previous rows may be flushed
    IStreamingRow AppendRow(int index);           // explicit index; must be > last index written

    NPOI.XSSF.Streaming.SXSSFSheet Underlying { get; }
}

public interface IStreamingRow
{
    int Index { get; }                            // 1-based
    IStreamingSheet Sheet { get; }
    IStreamingCell Cell(int col);                 // 1-based — see I-49
    IStreamingCell this[int col] { get; }
    IStreamingCell this[string column] { get; }   // row["B"]
    IStreamingRow Set(int col, string|double|decimal|int|long|bool|DateTime value);  // fluent (one overload per scalar)
    void Flush();                                 // explicit; auto-flushed on next AppendRow
}

// I-49 (added 2026-05-16): IStreamingRow.Cell returns IStreamingCell,
// not ICell. NPOI's SXSSFCell does not inherit from XSSFCell, so the
// ICell.Underlying contract (returns XSSFCell) cannot be honored on
// streaming cells. IStreamingCell is the narrower surface — value
// setters + Style/NumberFormat + Kind — with the escape hatch on
// IStreamingSheet.Underlying instead of per-cell.
public interface IStreamingCell
{
    string Address { get; }
    int RowIndex { get; } int ColumnIndex { get; }
    CellKind Kind { get; }
    void SetString(string value);
    void SetNumber(double|decimal|int|long value);
    void SetBool(bool value);
    void SetDate(DateTime value);
    void SetFormula(string formula);
    IStreamingCell Style(CellStyle style);
    IStreamingCell NumberFormat(string format);
}
```

### 6.4 Sheet

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
    IRow AppendRow();                             // append after last written row; row 1 if empty
    IRow Row(int index);                          // sheet.Row(1) == first row
    IColumn Column(int index);                    // sheet.Column(1) == "A"
    IColumn Column(string letter);                // sheet.Column("B")

    IEnumerable<IRow> Rows();                     // populated rows only
    IEnumerable<IRow> Rows(int startInclusive, int endExclusive);

    // (Typed row I/O is provided via source-generated extension methods —
    //  see §6.8. No methods on ISheet itself; AOT-safe by construction.)

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

### 6.5 Cell (fluent)

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
    ICell SetNumber(decimal value);               // rounded via (double)value; see §7.4 on precision
    ICell SetNumber(int value);                   // added 2026-05-16 — resolves the literal-`42` overload ambiguity
    ICell SetNumber(long value);                  // exact integer up to ±2^53; loses precision above
    ICell SetDate(DateTime value);
    ICell SetDate(DateOnly value);
    ICell SetTime(TimeOnly value);                // stored as Excel time fraction; needs time format string
    ICell SetDuration(TimeSpan value);            // stored as Excel time fraction; use [h]:mm:ss for elapsed
    ICell SetBool(bool value);
    ICell SetFormula(string formula);
    ICell Clear();

    // Typed reads
    string? GetString();
    double? GetNumber();
    DateTime? GetDate();
    TimeOnly? GetTime();
    TimeSpan? GetDuration();
    bool? GetBool();
    string? GetFormula();
    CellError? GetError();
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
public enum CellError { Null, DivByZero, Value, Ref, Name, Num, NotAvailable, GettingData }
[Flags] public enum BorderSide { None=0, Top=1, Right=2, Bottom=4, Left=8, All=15 }
public enum BorderStyle { None, Thin, Medium, Thick, Double, Dashed, Dotted }
public enum HAlign { General, Left, Center, Right, Fill, Justify }
public enum VAlign { Top, Center, Bottom }
public enum UnderlineStyle { None, Single, Double, SingleAccounting, DoubleAccounting }
```

### 6.6 Row / Column / Range

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
    ICell Cell(int col);                          // method form of this[int]
    IRow ForEachPopulated(Action<ICell> apply);   // visits populated cells only

    // Fluent setters — write a value to the column and return this row for chaining.
    // Pattern is sheet.AppendRow().Set(1, region).Set(2, revenue).Set(3, margin)...
    // Added to the design 2026-05-16 after the TabularExport cookbook recipe
    // surfaced the ergonomic motivation.
    IRow Set(int col, string value);
    IRow Set(int col, double value);
    IRow Set(int col, decimal value);
    IRow Set(int col, int value);
    IRow Set(int col, long value);
    IRow Set(int col, bool value);
    IRow Set(int col, DateTime value);
    IRow Set(int col, DateOnly value);
    IRow Set(int col, TimeOnly value);
    IRow Set(int col, TimeSpan value);

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

    IColumn AutoSize();                           // throws MissingFontException on headless Linux without fonts (I3)
    IColumn ForEachPopulated(Action<ICell> apply);
    IColumn SetDefaultStyle(CellStyle style);     // applies as the column-level default; new cells in this column inherit
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

    IEnumerable<ICell> EnumerateAll();            // dense; lazy (yield-based); materializes empty cells on demand

    // Dense operations — materialize every cell in the rectangle:
    IRange Value(object? value);                  // sets every cell
    IRange Apply(CellStyle style);                // sets style on every cell

    // Sparse iteration — visits only currently-populated cells:
    IRange ForEachPopulated(Action<ICell> apply);

    IRange Merge();
    IRange ClearContents();
}
```

### 6.7 Color

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

### 6.8 Style value object

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

### 6.8.1 Rich text (multi-run formatted strings) — I-50

```csharp
public sealed record RichTextStyle
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public UnderlineStyle? Underline { get; init; }
    public string? FontName { get; init; }
    public double? FontSize { get; init; }
    public Color? Color { get; init; }
    public static RichTextStyle Default { get; }
}

public sealed record RichTextRun(string Text, RichTextStyle Style)
{
    public RichTextRun(string text);  // RichTextStyle.Default
}

public sealed record RichText
{
    public RichText(IReadOnlyList<RichTextRun> runs);
    public RichText(params RichTextRun[] runs);
    public IReadOnlyList<RichTextRun> Runs { get; }
    public string PlainText { get; }   // string.Concat(Runs.Select(r => r.Text))
}

// On ICell:
void SetRichText(RichText value);
RichText? GetRichText();              // null for plain SetString cells
```

**I-50 (added 2026-05-22):** Rich text is modeled as immutable value records mirroring `CellStyle`'s shape, with the per-run style restricted to **font-only** axes. Excel's OOXML run model has no per-run fills, borders, alignment, or number format — exposing the full `CellStyle` on a run would silently drop those axes. Cell-level visual style remains routed through `ICell.Style(CellStyle)`; per-run typography is layered via `SetRichText`.

`GetRichText()` returns non-null only when the cell carries explicit formatting runs (`NumFormattingRuns > 0` in OOXML terms). A cell set via `SetString` returns `null` from `GetRichText()` even though OOXML stores every string cell as an `XSSFRichTextString` internally — the distinction surfaces "this string has run-level formatting" vs "this is a plain string".

`CellKind` stays `String` for rich-text cells (no new enum value). `GetString()` continues to return the concatenated plain text.

**Streaming write is intentionally absent.** `IStreamingCell` does **not** carry `SetRichText`. NPOI's SXSSF `SheetDataWriter` (NPOI 2.7.x) constructs a fresh `XSSFRichTextString` from `cell.StringCellValue` at flush time — dropping any in-memory formatting runs. Per decision #7 (type-honesty for streaming), the absence of the method on `IStreamingCell` mirrors the absence of the capability, rather than silently degrading or throwing at runtime. A future NPOI 3.x evaluation may revisit this.

The font pool used for cell-level styles (decision #4) is reused for run fonts — runs with identical font properties share one `IFont` across the workbook.

### 6.9 Typed mapping (source-generated extension methods)

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

For each `[Worksheet]`-annotated type `T`, the source generator emits a static class with these **extension methods on `ISheet`**:

```csharp
public static class SalesRecord_SheetExtensions   // generated, internal-by-default
{
    public static void AddRow(this ISheet sheet, SalesRecord record);
    public static void AddRows(this ISheet sheet, IEnumerable<SalesRecord> records);
    public static IEnumerable<SalesRecord> ReadRows(this ISheet sheet, int? headerRow = 1);
    //   headerRow: 1 (default) = row 1 is header; null = no header (positional by [Column(Order = N)])
}
```

Consequences:

- **No runtime reflection** for typed mapping. AOT and trim-safe.
- **No fallback path.** A type without `[Worksheet]` cannot be passed to `AddRow`/`AddRows`/`ReadRows` — the extension method does not exist for that type, and the code does not compile. There is no exception to catch because the call site never reaches runtime.
- The same extension methods exist on `IStreamingSheet` for the streaming write path.

### 6.10 A1 / range parser grammar

The cell and range parser accepts these forms. All accepted forms are normalized to the canonical form on output (`ICell.Address`, `IRange.Address`).

| Input form          | Accepted in `["..."]`? | Accepted in `Range("...")`? | Canonical form          | Notes                                  |
|---------------------|------------------------|-----------------------------|-------------------------|----------------------------------------|
| `A1`                | yes                    | yes (as single-cell range)  | `A1`                    | Uppercase letters                      |
| `a1`                | yes                    | yes                         | `A1`                    | Case-insensitive                       |
| `$A$1` / `$A1` / `A$1` | yes                 | yes                         | `A1`                    | `$` stripped; absolute/relative not modeled in v1 |
| `A1:C10`            | n/a                    | yes                         | `A1:C10`                | Standard range                          |
| `A:A`               | n/a                    | yes                         | `A1:A1048576`           | Whole column expanded to Excel max     |
| `1:1`               | n/a                    | yes                         | `A1:XFD1`               | Whole row expanded                     |
| `Sheet1!A1`         | **no**                 | **no**                      | —                       | Sheet-qualified refs are only valid in named-range formulas |
| Anything else       | throws                 | throws                      | —                       | `InvalidCellAddressException`           |

> **Dense operations on whole-row / whole-column ranges.** `Range("1:1").Value(0)` materializes ~16K cells; `Range("A:A").Value(0)` materializes ~1M. Prefer `ForEachPopulated` for read-side iteration over whole lines. For deliberate fills of a whole row/column, the dense `Value` call is allowed but the cost is the caller's to own.

### 6.11 Built-in number formats

```csharp
public static class NumberFormats
{
    public const string General           = "General";
    public const string Text              = "@";

    public const string Integer           = "0";
    public const string Number            = "#,##0";
    public const string NumberTwo         = "#,##0.00";
    public const string Scientific        = "0.00E+00";

    public const string Percent           = "0%";
    public const string PercentTwo        = "0.00%";

    public const string Currency          = "$#,##0.00";
    public const string CurrencyNoSymbol  = "#,##0.00";
    public const string Accounting        = "$#,##0.00;[Red]-$#,##0.00";  // negative in red

    public const string Date              = "yyyy-mm-dd";
    public const string DateTime          = "yyyy-mm-dd hh:mm:ss";
    public const string Time              = "hh:mm:ss";
    public const string Duration          = "[h]:mm:ss";       // elapsed; renders 25:00:00 etc.
}
```

This set is frozen in v1.0. Additions in later releases require an explicit decision row.

### 6.12 Source generator: diagnostic catalog

The `[Worksheet]` source generator emits diagnostics under the `NXLS` prefix. v1.0 catalog:

| ID            | Severity | Meaning                                                                                  |
|---------------|----------|------------------------------------------------------------------------------------------|
| `NXLS0001` | Error    | `[Worksheet]` type has duplicate `[Column]` orders                                       |
| `NXLS0002` | Error    | `[Worksheet]` type has no public parameterless constructor and no designated constructor. **Records with primary constructors satisfy this rule** — the primary constructor *is* the designated ctor |
| `NXLS0003` | Error    | `[Column]` attribute references a `Format` string that fails the parser smoke check       |
| `NXLS0004` | Warning  | `[Worksheet]` type has properties not marked `[Column]` or `[Ignore]` (silent skip is ambiguous) |
| `NXLS0005` | Error    | `[Worksheet]` type is not `partial`. Required so the generator can emit nested helper types alongside it in later versions without breaking source layout; future-proofing decision recorded in v1.0 |
| `NXLS0006` | Error    | Property type has no built-in converter (v1.0). In v1.1 this expands to: "and no `[ColumnConverter]` attribute" once custom converters ship |

Generator behavior summary:

- `IIncrementalGenerator` (modern Roslyn API).
- **Scans only the current compilation.** This is a load-bearing constraint with downstream user impact, called out here and again in the public XML documentation on the `[Worksheet]` attribute itself. `[Worksheet]`-annotated types defined in a *referenced assembly* (`MyShared.dll` consumed by `MyApp.csproj`) are **invisible** to the generator running in `MyApp.csproj` — the extension methods are not emitted, calls do not compile. Each assembly that defines `[Worksheet]` types must add the `NetXlsx` package itself so the generator runs against that compilation. This matches the scoping rule used by `System.Text.Json`'s source generator. Common failure mode: a consumer puts shared records in a "Domain" library expecting them to "just work" in the calling app — they do not. The fix is to add `NetXlsx` to the Domain library too.
- Emits source files visible under `obj/Generated/` when `<EmitCompilerGeneratedFiles>true</EmitCompilerGeneratedFiles>` is set in the consuming project (useful for debugging).
- Generated extension classes are `internal` by default; consumers may opt into `public` via `[Worksheet(Visibility = Visibility.Public)]`.

### 6.13 Exception hierarchy

```csharp
public class WorkbookException : Exception { ... }
public sealed class InvalidCellAddressException : WorkbookException { ... }
public sealed class SheetNameException : WorkbookException { ... }
public sealed class StyleBudgetExceededException : WorkbookException { ... }
public sealed class MalformedFileException : WorkbookException { ... }
public sealed class ResourceLimitExceededException : WorkbookException { ... }
public sealed class FormulaException : WorkbookException { ... }
public sealed class MissingFontException : WorkbookException { ... }   // I3 — AutoSize on headless Linux
```

## 7. Behavioral notes

### 7.1 Async semantics

NPOI is synchronous. `SaveAsync`/`OpenAsync` use `Task.Run` to offload mixed CPU+I/O work to the thread pool. We do *not* return `Task.FromResult` over trivially-synchronous work; if a method has no async path, it stays synchronous.

Callers in ASP.NET / Blazor contexts should note: `SaveAsync` consumes a thread-pool thread for the duration of the save. If you serialize many workbooks concurrently, throttle.

**Cancellation.** `CancellationToken` is honored only at offload boundaries — before the work is dispatched to the thread pool, and after NPOI returns. Mid-operation cancellation is not supported because NPOI itself is not cancellation-aware. A token cancelled while NPOI is mid-save will not interrupt the save; it will surface as `OperationCanceledException` only on the next async boundary.

### 7.2 Culture handling

Excel stores numbers as IEEE doubles regardless of the spreadsheet's locale; reading via `GetNumber()` / `GetValue<double>()` is culture-invariant by construction.

**Number-format strings are pass-through.** When the caller writes `cell.NumberFormat("#,##0.00")`, those exact bytes are stored. The facade does not transform, localize, or interpret format strings on write. Excel renders the cell according to *its own* locale settings at display time, not anything NetXlsx does.

`WorkbookOptions.DisplayCulture` therefore affects exactly one thing: how `GetString()` against a number- or date-formatted cell produces a displayed-string representation on the read path, when the caller asks for a string rather than the raw typed value.

Reading raw values: always invariant. Reading displayed strings: `DisplayCulture`-aware. Writing format strings: pass-through.

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

### 7.4 `decimal` write semantics

`.xlsx` cells store numeric values as IEEE-754 `double` only — there is no native decimal type in the format. `SetNumber(decimal value)` therefore rounds via `(double)value` before storage:

- Values within `double`'s 15–17 significant-digit range round-trip exactly.
- Values outside that range silently lose precision.
- Values outside `double`'s magnitude range (extremely large/small) clamp to `±double.MaxValue`.

If exact textual preservation is required (currency with > 15 significant digits, identifiers that look numeric, etc.), use `SetString` and apply a number-format string for display only. The facade will not silently switch to text representation under the hood — the type the caller writes is the type stored.

### 7.5 Round-trip style preservation

When `Workbook.Open` reads a file, existing cell styles are wrapped by reference and **not** imported into the dedup cache. The cache acts only on styles created through the facade's fluent API after open.

Consequence: opening and immediately re-saving a file produces no changes to the `xl/styles.xml` part. This is intentional — it keeps the facade diff-friendly when used in pipelines that round-trip workbooks for inspection or minor edits.

Mixing reads and new fluent styling on the same cell is well-defined: the dedup cache hashes against the resulting `CellStyle` value record, so a style "originally from open" plus a fluent `.Bold()` produces a new (cached) style and leaves the original untouched.

### 7.6 Excel hard-limit enforcement

Writes that would exceed Excel's hard format limits throw `ResourceLimitExceededException` at the call site. No silent truncation, no auto-rollover.

| Limit                          | Default                | Option                            |
|--------------------------------|------------------------|-----------------------------------|
| Rows per sheet                 | 1,048,576              | `WorkbookOptions.MaxRowsPerSheet` |
| Columns per sheet              | 16,384 (`XFD`)         | `WorkbookOptions.MaxColsPerSheet` |
| Cell text length (characters)  | 32,767                 | `WorkbookOptions.MaxCellTextLength` |
| Sheet name length (chars)      | 31                     | (validated by `Workbook.IsValidSheetName` / `SanitizeSheetName`) |
| Sheet count (read-side safety) | 1,000                  | `WorkbookOptions.ReadMaxSheets`   |
| Uncompressed size (read)       | 256 MB                 | `WorkbookOptions.ReadMaxUncompressedBytes` |

Defaults match Excel's hard caps; consumers can lower them but cannot raise them above what Excel itself accepts.

### 7.7 Round-trip preservation of unknown XML parts

Workbooks routinely contain parts NetXlsx does not model in v1: pivot caches, conditional formatting, custom XML, threaded comments, vendor `<ext>` elements, and any future Microsoft additions. **`Open → Save` preserves every such part verbatim, by reference, via the OPC package layer.** The facade never reaches into parts it does not understand.

Consequences:

- A round-trip with no fluent modifications produces a byte-equivalent file (modulo deterministic whitespace / part-ordering normalization from the OPC writer).
- Workbooks containing v2 features (charts, conditional formatting) can be opened and minimally edited in v1.0 without losing those features.
- A dedicated round-trip preservation test in the golden-file corpus (`OpenEditSave` recipe — see §8.1) is a v1.0 ship-blocker. It opens a workbook containing every part type we don't model, modifies one cell, saves, and asserts the unmodeled parts are bit-identical.

If `Open → modify → Save` *does* drop an unknown part, that is a bug, not a feature.

### 7.8 Formula values are not pre-computed

When a formula is written via `cell.SetFormula("=A1+B1")`, the cell is stored with no cached value. Excel, LibreOffice, Google Sheets, and any other competent consumer recalculates on open.

NetXlsx does **not** invoke NPOI's formula evaluator on save. Pre-computation is slow on large workbooks, fragile for any formula touching newer functions, and produces results that may diverge from Excel's own evaluation. Letting the consumer recompute is both faster and more correct.

The trade: a file briefly viewed by a tool that does not recalculate (rare) will show blank values for formula cells until the user triggers a recalc. We consider this an acceptable cost.

### 7.9 Time and duration handling

Excel has no native time type. `SetTime(TimeOnly)` stores the time-of-day as a fraction of a day (`0.5` = noon). `SetDuration(TimeSpan)` stores the same way; durations beyond 24h require an elapsed-time format string like `[h]:mm:ss`.

Format strings are the caller's responsibility — without one, the cell shows as a fractional number. The cookbook ships a `TimeAndDuration` recipe demonstrating the common patterns.

`GetTime` and `GetDuration` reverse the conversion; they accept any numeric cell and interpret it as a time fraction. Callers must apply judgment about whether the cell semantically represents a time.

**Negative durations** (`TimeSpan` with `< TimeSpan.Zero`) throw `ArgumentOutOfRangeException` from `SetDuration`. Excel cannot render negative time natively; silently storing a negative fraction would produce a file that displays incorrectly downstream.

### 7.10 `DateTime.Kind` and type-conversion semantics

`SetDate(DateTime)` stores the value verbatim — no timezone conversion is performed, regardless of `Kind`. Excel has no native timezone concept; converting under the hood would silently lose information. If the caller has timezone semantics, they convert before writing.

`GetDate()` returns a `DateTime` with `Kind = Unspecified` on every cell. There is no metadata in the file to preserve the original `Kind`.

`GetString()` rendering rules (independent of `DisplayCulture` for these types):

| Cell content        | `GetString()` returns                                  |
|---------------------|--------------------------------------------------------|
| Empty               | `""`                                                   |
| String              | The stored string verbatim                             |
| Number              | Invariant-culture string of the value; respects format string if applied via `DisplayCulture` |
| Date                | Display-formatted per the cell's number format and `DisplayCulture` |
| Boolean             | `"TRUE"` or `"FALSE"` (invariant; never localized)     |
| Formula             | The cached formula result as a string (per the rules above for its underlying type) |
| Error               | The error code as text: `"#DIV/0!"`, `"#REF!"`, etc.   |

`GetValue<T>()` returns `null` for `T?` (or `default(T)` for non-nullable `T`) when the cell's actual type cannot be converted to `T`. It does not throw on type mismatch. Examples:

- `cell.GetValue<int>()` on a string cell → `0` (default).
- `cell.GetValue<int?>()` on a string cell → `null`.
- `cell.GetValue<string>()` on any cell → the `GetString()` rendering above.
- `cell.GetValue<DateTime?>()` on a numeric cell with date formatting → the date.

## 8. Repository layout

```
NetXlsx/
├─ src/
│  ├─ NetXlsx/                  # main library — NetXlsx namespace
│  │   └─ Streaming/               # IStreamingWorkbook etc. — NetXlsx.Streaming namespace
│  ├─ NetXlsx.SourceGen/        # source generator for [Worksheet] mapping
│  └─ NetXlsx.Analyzers/        # Roslyn analyzers (v2+) — NXLS#### diagnostic IDs
├─ tests/
│  ├─ NetXlsx.Tests/            # xUnit unit tests
│  ├─ NetXlsx.GoldenFiles/      # reference workbooks + round-trip tests
│  └─ NetXlsx.PublicApi/        # public-API snapshot tests
├─ benchmarks/
│  └─ NetXlsx.Benchmarks/       # BenchmarkDotNet vs NPOI/EPPlus/ClosedXML
├─ samples/
│  └─ NetXlsx.Cookbook/         # runnable recipes (see §8.1)
├─ docs/
│  ├─ design.md                    # this document
│  ├─ roadmap.md                   # feature roadmap + binary scope per release
│  └─ npoi-workarounds.md          # bug catalogue for NPOI issues we work around
├─ build/
│  ├─ build.ps1                    # local + CI entry point (Windows)
│  └─ build.sh                     # local + CI entry point (Linux/macOS)
├─ .github/workflows/              # GitHub Actions CI + release pipelines
├─ .editorconfig                   # code style (whole repo)
├─ Directory.Build.props           # shared MSBuild settings (TFM, nullable, deterministic, signing)
├─ Directory.Packages.props        # central package management (single NPOI pin)
├─ nuget.config                    # public NuGet feed
├─ CODEOWNERS                      # PR review routing
├─ CHANGELOG.md                    # Keep-a-Changelog format
├─ LICENSE                         # MIT
├─ NetXlsx.sln
└─ README.md
```

### 8.1 Minimum cookbook recipes for v1.0

Each recipe is a single runnable program under `samples/NetXlsx.Cookbook/` that demonstrates one capability cleanly. v1.0 ships with at least these:

1. **HelloWorkbook** — Create, add a sheet, write a few cells, save.
2. **TabularExport** — Write 10k records from a list, with a header row, frozen header, and column widths.
3. **TypedExport** — Same as above using `[Worksheet]` + source-gen extension methods.
4. **TypedImport** — Read a workbook into a typed record sequence.
5. **StyledReport** — Demonstrate the fluent style API: bold headers, currency formatting, conditional cell coloring.
6. **Formulas** — Write formulas referencing other cells; demonstrate that Excel computes on open.
7. **MultiSheet** — Three sheets (summary, data, lookup) with named ranges and cross-sheet formulas.
8. **HyperlinksAndComments** — Annotated cells with comments and external hyperlinks.
9. **StreamingMillionRows** — `Workbook.CreateStreaming()` to write a million rows under the perf budget.
10. **NPOIEscapeHatch** — Use `.Underlying` to do something the facade doesn't cover (e.g., set a print area).
11. **OpenEditSave** — Open an existing workbook, modify a few cells, save — demonstrate the no-churn styles guarantee (§7.5) *and* the unknown-parts preservation guarantee (§7.7).
12. **TimeAndDuration** — Demonstrate `SetTime(TimeOnly)` and `SetDuration(TimeSpan)` with appropriate format strings, including elapsed-time formatting via `[h]:mm:ss`.
13. **CellErrors** — Read a workbook containing formula errors; classify them via `GetError()` and the `CellError` enum.

Each recipe is also a golden-file test: the produced workbook is compared byte-by-byte (or structurally for non-deterministic parts) against a reference under `tests/NetXlsx.GoldenFiles/`.

## 9. Engineering substrate

The decisions in §3 govern the API and contracts. The decisions below govern the project itself — toolchain, distribution, hygiene. They are made up-front so that an agent (or human) can scaffold the v1.0 project without further pause.

### 9.1 Toolchain

| #  | Decision                       | Choice                                                                       | Rationale                                                          |
|----|--------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------|
| S1 | Test framework                 | xUnit                                                                        | Idiomatic for modern .NET libraries; parallel by default; ubiquitous tooling support |
| S2 | Test runner + assertion library | Classic VSTest host (`Microsoft.NET.Test.Sdk` + `xunit.runner.visualstudio`); assertions via xUnit's `Assert` plus `FluentAssertions v6` for readable golden-file diffs. Migration to `Microsoft.Testing.Platform` is deferred — revisit when MTP is the default in the .NET SDK | VSTest is the mature, ubiquitous path today; MTP is the strategic future. FluentAssertions pinned to v6 to avoid v7's license change |
| S3 | Benchmark framework            | BenchmarkDotNet                                                              | Industry standard; produces statistically valid measurements        |
| S4 | Test naming convention         | `Method_Scenario_Expected`                                                   | Concise; matches Microsoft guidance; greps cleanly                  |
| S5 | Code-style file                | `.editorconfig` at repo root, enforced in CI                                 | One source of truth for whitespace, ordering, naming               |
| S6 | Public-API surface tracking    | `Microsoft.CodeAnalysis.PublicApiAnalyzers` (Shipped/Unshipped text files)   | Forces deliberate API changes; CI fails on undeclared additions     |
| S7 | Source generator hosting       | In-repo project `src/NetXlsx.SourceGen/`; ships in the same NuGet package | One package; no separate install dance for consumers               |

### 9.2 Distribution

| #   | Decision                       | Choice                                                                       | Rationale                                                          |
|-----|--------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------|
| S8  | NPOI version pin               | Exact pin in `Directory.Packages.props`; bump deliberately via PR             | NPOI is a public dependency (#24); we control the bump cadence     |
| S9  | License                        | MIT                                                     | Internal use; can be re-licensed (e.g., Apache-2.0) if open-sourced later — no copyleft contributions block this |
| S10 | NuGet feed                     | public NuGet feed (URL set in `nuget.config` at scaffold time)                | Public distribution via nuget.org                                          |
| S11 | Versioning                     | MinVer (git-tag-driven SemVer)                                               | Zero-config; SemVer comes from `git tag`; no version state in source files |
| S12 | Strong-name signing            | Yes; key in repo (`netxlsx.snk`)                                          | Required for some legacy consumers; cheap to add               |
| S13 | Symbol packages                | `.snupkg` published alongside `.nupkg`                                       | Debugging across package boundaries                                |
| S14 | Source Link                    | Enabled (`Microsoft.SourceLink.GitHub` or equivalent for github.com/jkindrix)  | Step-into-source from consumers                                    |
| S15 | Deterministic builds           | Enabled (`<Deterministic>true</Deterministic>`, `ContinuousIntegrationBuild=true` on CI) | Reproducible binaries; verifiable across machines             |
| S16 | Diagnostic ID scheme           | Unified `NXLS<NNNN>` 4-digit format with documented ranges per category: `0001-0099` source-generator diagnostics, `0100-0199` MSBuild build-time guards, `0200-0299` reserved for Roslyn analyzers (v2+), `0300+` reserved | Identifiable; no clashes with common Microsoft prefixes; single uniform format avoids the dual-scheme renumbering churn the alternative (`NXLSAOT001` vs `NXLS0001`) would invite |

### 9.3 CI / process

| #   | Decision                       | Choice                                                                       | Rationale                                                          |
|-----|--------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------|
| S17 | CI platform                    | GitHub Actions; workflows under `.github/workflows/` (CI on push/PR, release on tag) | GH Actions is the default for OSS hosted on GitHub; zero infra cost |
| S18 | Local + CI build entry point   | `build/build.ps1` (Windows) and `build/build.sh` (Unix); both call the same MSBuild targets | One command runs the whole build locally; CI calls the same script |
| S19 | Branching                      | Trunk-based (`main`); short-lived feature branches; no long-running releases | Modern default; matches v1.0 cadence                               |
| S20 | PR policy                      | All changes via PR; at least one approval; all CI checks must pass            | Minimum bar; can tighten later                                     |
| S21 | CHANGELOG format               | Keep-a-Changelog                                                             | Conventional; readable; tool-friendly                              |
| S22 | CODEOWNERS                     | Routes all changes to a small named maintainer group                          | Avoids PR review going to /dev/null                                |
| S23 | Dependency updates             | Renovate, enabled post-v1.0 (deferred — see §9.5)                            | Useful but not blocking; intentional bumps via PR in the meantime  |

### 9.4 Documentation

| #   | Decision                       | Choice                                                                       | Rationale                                                          |
|-----|--------------------------------|------------------------------------------------------------------------------|--------------------------------------------------------------------|
| S24 | XML docs                       | Mandatory on every public symbol; CI fails on missing                         | Quality gate from §5                                               |
| S25 | API documentation site         | Deferred to v1.1 (DocFX); v1.0 ships in-repo markdown only                   | Internal use first; site can come later                            |
| S26 | NPOI workaround catalog        | `docs/npoi-workarounds.md`, populated as workarounds are discovered           | Surfaces NPOI bugs we route around; informs future upstream PRs    |
| S27 | AOT/trim consumer build guard  | `buildTransitive/NetXlsx.targets` shipped in nupkg emits MSBuild errors `NXLS0100` (`PublishAot`) and `NXLS0101` (`PublishTrimmed`). Latter has `<NetXlsxAllowTrimmed>true</…>` escape hatch. IDs fall in the `0100-0199` MSBuild-guard range per S16 | Enforce decision I2 at build time rather than relying on runtime failure; added 2026-05-16, IDs unified 2026-05-16 |
| S28 | Analyzer suppressions          | `.editorconfig` suppresses `CA1716` (Set vs VB keyword), `CA1720` (CellKind.String), `RS0026` (intentional optional-param overloads). Each suppression has an inline rationale comment | C#-only internal library; design uses identifiers/patterns these rules consider ambiguous in other languages |
| S29 | Date-time default style cache  | `XssfWorkbook` holds a lazy per-workbook cache of four format-only `ICellStyle` instances (date / datetime / time / duration). NOT the full §4 style pool — interim until the styling-API slice lands. Composes cleanly: the four cached styles will register as dedup entries once the pool exists | Avoid per-cell style allocation in the date/time slice without blocking on the full pool |

### 9.5 Deferred (captured; not addressed in v1.0)

These items are intentionally not addressed in v1.0. They are listed so a future maintainer can find them and address them when the time is right.

| Item                                          | Target           | Reason for deferral                                       |
|-----------------------------------------------|------------------|-----------------------------------------------------------|
| DocFX documentation site                       | v1.1             | In-repo markdown is sufficient for internal launch        |
| Renovate / Dependabot                          | post-v1.0        | Manual NPOI bumps are deliberate while API stabilizes     |
| Plugin / extension points (custom converters)  | v2 (pending demand) | No concrete user ask yet; YAGNI                        |
| Dynamic-array / spill-range formulas           | v3 or never      | Excel 365 feature; rare in our consumer base              |
| Memory-mapped file reads                       | v3 or never      | Optimization without a current bottleneck                 |
| Workbook encryption (file-level password)      | v3               | Already in feature roadmap                                |
| External-reference resolution                  | never            | Preserve as opaque; do not chase                          |
| Pivot table *reading*                          | never (see roadmap) | Out of scope explicitly                                |

### 9.6 What an agent now knows

With §3, §6, §7, §8, and §9 in hand, an agent has:

- The full public API surface for v1.0.
- The behavioral contracts for every edge case we've identified.
- The exact set of files to scaffold, the namespaces they live in, and the toolchain to use.
- The CI platform, build entry points, packaging policy, and versioning mechanism.
- The list of pre-implementation spikes that gate locking the design (in `roadmap.md`).
- The minimum cookbook recipe set and that each is a golden-file test.

Open items requiring user input at scaffold time (and only at scaffold time):

- The public NuGet feed URL (goes in `nuget.config`).
- The github.com/jkindrix URL (goes in Source Link config).
- The CODEOWNERS group identifiers.
- The strong-name key bytes (or the decision to generate one at scaffold).

Everything else is decided.

## 10. Definition of "extraordinary" for this project

- Code that's a pleasure to read.
- Abstractions that make complex things (style management, typed mapping) simple.
- Error messages that help the caller recover (`InvalidCellAddressException` says *why* the address was invalid, what was expected, and what they passed).
- Performance per the targets in §4 — `< 10%` overhead vs raw NPOI on typical workloads (≤ 100 distinct cell styles); `< 30%` in worst-case style-dedup-heavy scenarios. The two-tier target reflects honest measurement, not handwaving.
- Documentation that clarifies rather than restates.
- The escape hatch is honest: nothing the facade does is uncovertible to a raw-NPOI equivalent.
