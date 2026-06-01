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

### 6.2.1 Workbook protection — I-54

```csharp
public sealed record WorkbookProtection
{
    public bool Structure { get; init; }
    public bool Windows { get; init; }
    public bool Revision { get; init; }
    public static WorkbookProtection Default { get; }         // all false
    public static WorkbookProtection LockStructure { get; }   // Structure-only
}

// On IWorkbook:
void Protect(WorkbookProtection? options = null);
void Unprotect();
bool IsProtected { get; }
```

**I-54 (added 2026-05-22):** Workbook-level protection guards the workbook's structure (sheet add/delete/rename/reorder/hide), windows (Excel 2007-era window chrome — largely defunct in modern Excel), and revisions (legacy shared-workbook tracking).

`Protect()` with no arguments applies `WorkbookProtection.LockStructure` — the common use case is "stop the user from accidentally adding or deleting sheets." Pass an explicit options record to lock additional facets.

Same security caveat as sheet protection (I-53): this is a UX guard, not security. NPOI 2.7.3 does not expose workbook-level password support directly; password protection requires reaching through `.Underlying` and manipulating the `CT_WorkbookProtection` element. v1.1 ships the unprotected-by-default flag flip; password support is deferred.

### 6.2.2 Named / reusable styles — I-57

```csharp
// On IWorkbook:
void RegisterStyle(string name, CellStyle style);
CellStyle? GetRegisteredStyle(string name);   // null if unknown
IReadOnlyCollection<string> RegisteredStyleNames { get; }

// On ICell:
ICell ApplyNamedStyle(string name);

// On IRange:
IRange ApplyNamedStyle(string name);
```

**I-57 (added 2026-05-22):** A workbook-scoped name registry for `CellStyle` values. Register once with a name; apply by name on cells or ranges. Backed by the existing style-pool dedup (decision #4) — equal `CellStyle` instances still share one underlying NPOI `ICellStyle`, so naming is purely a caller-side convenience.

**v1.1 named styles are an in-process convenience, not OOXML named-style table entries.** When a workbook is saved and reopened via `Workbook.Open`, the per-cell style is preserved (via style-pool dedup), but the name → style map is not rehydrated — Excel itself never saw the names. Real OOXML named-style table integration (so styles persist by name across open/save) is deferred to v1.2.

Names are case-insensitive. Re-registering an existing name replaces the definition.

### 6.2.3 Strict concurrency detection — I-59

```csharp
// On WorkbookOptions:
public bool StrictConcurrencyDetection { get; init; }   // default false
```

**I-59 (added 2026-05-22):** Opt-in real-lock mode replacing the opportunistic reentry counter (decision #43) for callers who want a hard guarantee against silent corruption from concurrent threads.

When `false` (default): the workbook uses an `Interlocked.CompareExchange`-based reentry counter on `EnterMutation`. Concurrent mutation throws `InvalidOperationException` *opportunistically* — the check has a brief gap during which two threads can both observe a "no mutation in progress" state and proceed. Same-thread reentrancy is also rejected.

When `true`: `EnterMutation` takes a real per-workbook `Monitor` lock. Concurrent mutations serialize cleanly (no exception); reentrant mutations on the same thread are permitted (Monitor is reentrant). Trades some throughput for a strong "you cannot silently corrupt this workbook even if you ignore the thread-safety docs" guarantee.

The default-mode exception message explicitly references the option so callers discover it when they hit a contention throw.

Strict mode does not make the workbook thread-safe for reads — concurrent reads of any kind remain undefined. The lock is for mutations only.

### 6.2.4 Style-pool diagnostics — I-61

```csharp
public readonly struct StylePoolDiagnostics
{
    public int StyleHitCount { get; }
    public int StyleMissCount { get; }
    public int FontHitCount { get; }
    public int FontMissCount { get; }
    public int UniqueStyles { get; }
    public int UniqueFonts { get; }
    public double StyleDedupRatio { get; }   // hits / (hits + misses); 0 when no lookups
    public double FontDedupRatio { get; }
}

// On IWorkbook:
StylePoolDiagnostics GetStylePoolDiagnostics();
```

**I-61 (added 2026-05-22):** Read-only operational visibility over the workbook's `CellStyle` + font dedup pools (decision #4). Surfaces "is the dedup actually working in production?" as a query, useful for ops sanity checks ("we have 50,000 cells; do we have 50 unique styles or 50,000?") and as test-time assertion material.

The counters track every call site that goes through `CellStylePool.GetOrCreate` and `GetOrCreateFont` — both the explicit `ICell.Style(CellStyle)` path and the internal default-style-on-write paths (e.g., `SetDate` applying the workbook's date style).

The returned struct is a **snapshot** by value. Calling `GetStylePoolDiagnostics` does not allocate, and the snapshot does not update after capture. A subsequent mutation requires a new call.

### 6.2.5 Workbook password protection — I-65 (v1.2)

```csharp
// On IWorkbook:
void Protect(WorkbookProtection? options = null);                          // v1.1 (I-54)
void ProtectWithPassword(string password, WorkbookProtection? options = null);  // v1.2 (I-65)
```

**I-65 (added 2026-05-22):** v1.1 slice 5 (decision I-54) shipped workbook-level structure / windows / revision locks without password support — NPOI 2.7.3 did not expose a workbook-level password helper. v1.2 closes the gap by writing the 16-bit XOR verifier directly into `CT_WorkbookProtection.workbookPassword`.

The verifier is computed via NPOI's `CryptoFunctions.CreateXorVerifier1(password)` (the public legacy-hash helper) and encoded as a 2-byte big-endian array. The OOXML `byte[]` field serializes via `XmlHelper.WriteAttribute(byte[])` as a hex string — matching Excel's expected format.

**Method-naming decision:** v1.2 uses a separate method (`ProtectWithPassword`) rather than an overload of `Protect`. Adding a second optional parameter to `Protect(WorkbookProtection? options = null)` would create a call-site ambiguity that the C# compiler refuses to resolve (both overloads would match `wb.Protect()`), and adding an overload without defaults would violate `RS0027` (the Roslyn "API with optional parameter(s) should have the most parameters amongst its public overloads" rule). A separately-named method is also more self-documenting at the call site:

```csharp
wb.Protect();                                  // structure-only lock, no password
wb.ProtectWithPassword("hunter2");             // structure + password
wb.ProtectWithPassword("hunter2", LockAll);    // explicit options + password
```

**Security caveat repeated from I-54:** the XOR-verifier algorithm is widely known to be brute-forceable. This is a UX guard ("Excel rejects the unprotect attempt if the password doesn't hash to the same verifier"), not real security.

`Unprotect` clears the structure / windows / revision flags but does **not** clear the `workbookPassword` byte array — once cleared of lock flags, the password is effectively dormant since no protected state references it. A subsequent `Protect` or `ProtectWithPassword` overwrites it.

### 6.2.6 OOXML named-style table integration — I-67 (v1.3)

**I-67 (added 2026-05-22):** v1.1 named styles (decision I-57) were an in-process convenience — `RegisterStyle` populated a `Dictionary<string, CellStyle>` but did not produce entries in OOXML's `cellStyleXfs` + `cellStyles` tables, so the name map was lost on `Workbook.Open`. v1.3 closes the loop.

**Write side (`RegisterStyle`):**

1. Allocate the visual style via the existing `CellStylePool.GetOrCreate` — gives a regular `cellXfs` entry whose font / fill / border indices we then mirror.
2. Build a fresh `CT_Xf` for `cellStyleXfs` referencing the same indices. `applyFont` / `applyFill` / `applyBorder` are set when the corresponding index is non-zero.
3. Append the `CT_Xf` to NPOI's internal `styleXfs` list via reflection (`StylesTable.PutCellStyleXf` is `internal` upstream — centralized in `NpoiInternals`). Direct manipulation of `CT_Stylesheet.cellStyleXfs.xf` fails because NPOI's save path overwrites that list with its internal copy at line 887 of `StylesTable.cs`.
4. Add (or update) a `CT_CellStyle` entry to `CT_Stylesheet.cellStyles.cellStyle` with `name` + `xfId` pointing at the new cellStyleXfs index.

**Read side (lazy on first access of `NamedStyles`):**

1. Walk `CT_Stylesheet.cellStyles.cellStyle`.
2. Skip the built-in `"Normal"` entry (Excel always creates one; surfacing it as a registered name would be noisy) and any entry with null/empty name.
3. For each entry, get the corresponding `CT_Xf` via `NpoiInternals.GetCellStyleXfAt`.
4. Materialize the CT_Xf as a regular `cellXfs` entry via `StylesTable.PutCellXf` (also `internal`; centralized in NpoiInternals). Wrap with `XSSFCellStyle`. Pass to the existing `CellStylePool.ReadFromNpoi(ICellStyle)` to produce a `CellStyle` value.
5. Populate the in-process `_namedStyles` dictionary directly (no re-write of the OOXML entry — it's already there).

Rehydration is triggered lazily from the `NamedStyles` property getter, which `GetRegisteredStyle` / `RegisteredStyleNames` / `RegisterStyle` all flow through.

**Limitations carried forward:**

- Excel's "Cell Styles" panel will show the named styles after open. Cells styled via `ApplyNamedStyle` do **not** carry a `xfId` reference to the named-style XF — they get an explicit cellXfs entry instead. The visual outcome is identical; the UI distinction (whether the cell shows as "Header" in the Cell Styles ribbon group) is not preserved across round-trip. Closing that gap fully would require the cellXfs entry to carry an `xfId` attribute pointing at the cellStyleXfs entry — deferred to a future slice if a user surfaces it as a need.
- Materializing the cellStyleXfs entry as a cellXfs entry at read time produces one duplicate cellXfs row per named style. Cost is bounded (one per named style, not one per styled cell); acceptable.

### 6.2.7 `.xlsm` macro-enabled passthrough — I-69

```csharp
// Entry point:
public static IWorkbook CreateMacroEnabled(WorkbookOptions? options = null);

// On IWorkbook:
bool IsMacroEnabled { get; }
```

**I-69 (added 2026-05-26):** NetXlsx is `.xlsx`-focused (decision #6) but many real-world workbooks carry VBA macros (`.xlsm`). The passthrough feature adds:

1. **`Workbook.CreateMacroEnabled(options?)`** — creates an `XSSFWorkbook(XSSFWorkbookType.XLSM)`. The resulting workbook uses the macro-enabled OOXML content type (`application/vnd.ms-excel.sheet.macroEnabled.main+xml`) so that VBA project parts can be added through `.Underlying` and survive round-trip.
2. **`IWorkbook.IsMacroEnabled`** — read-only boolean, delegates to NPOI's `XSSFWorkbook.IsMacroEnabled()`. True for workbooks created via `CreateMacroEnabled` or opened from `.xlsm` files.
3. **`Workbook.Open` transparently handles `.xlsm`.** No code change needed — NPOI's `XSSFWorkbook(Stream)` constructor detects the OOXML content type and preserves it. The open-path error messages were updated to mention `.xlsm` alongside `.xlsx`.

**What this is NOT:** NetXlsx does not read, write, inspect, or execute VBA. The macro content is passthrough only — VBA project parts survive `Open → Save` untouched via NPOI's OPC-part preservation (decision #44). Callers who need to inject or extract VBA reach through `.Underlying` or use a dedicated VBA library.

### 6.2.8 Grouping / outlining — I-71

```csharp
// On ISheet:
void GroupRows(int startRow, int endRow);
void UngroupRows(int startRow, int endRow);
void GroupColumns(int startCol, int endCol);
void UngroupColumns(int startCol, int endCol);
void SetRowGroupCollapsed(int row, bool collapsed);
```

**I-71 (added 2026-05-26):** Row and column grouping (Excel's "Group" / outline feature). Delegates to NPOI's `GroupRow`, `UngroupRow`, `GroupColumn`, `UngroupColumn`, and `SetRowGroupCollapsed`. All indices are 1-based per NetXlsx convention (converted to 0-based for NPOI).

Nested groups are supported — NPOI increments the outline level on each nested `GroupRow`/`GroupColumn` call, up to Excel's 7-level limit. `SetRowGroupCollapsed` toggles the collapsed state of a group whose summary row is at the given index.

**Column collapse** is not surfaced as a dedicated method in v1 — NPOI's column collapse behavior is less predictable than row collapse. Callers needing column collapse reach through `ISheet.Underlying`.

### 6.2.9 Conditional formatting — I-73

```csharp
public sealed class ConditionalFormat
{
    // Cell-value conditions
    public static ConditionalFormat CellValueGreaterThan(string value, CellStyle style);
    public static ConditionalFormat CellValueLessThan(string value, CellStyle style);
    public static ConditionalFormat CellValueBetween(string min, string max, CellStyle style);
    public static ConditionalFormat CellValueEqual(string value, CellStyle style);
    public static ConditionalFormat CellValueNotEqual(string value, CellStyle style);
    public static ConditionalFormat CellValueGreaterThanOrEqual(string value, CellStyle style);
    public static ConditionalFormat CellValueLessThanOrEqual(string value, CellStyle style);
    // Formula-based
    public static ConditionalFormat Formula(string formula, CellStyle style);
    // Color scales
    public static ConditionalFormat ColorScale(Color min, Color max);
    public static ConditionalFormat ColorScale(Color min, Color mid, Color max);
}

// On ISheet:
void AddConditionalFormatting(string a1Range, params ConditionalFormat[] rules);
int ConditionalFormattingCount { get; }
void RemoveConditionalFormatting(int index);
```

**I-73 (added 2026-05-26):** Exposes NPOI's `SheetConditionalFormatting` through a factory-based API matching the established `DataValidation` and `FilterCriteria` patterns. Three rule types:

1. **Cell-value conditions** — highlight cells based on comparison operators. The `style` parameter applies bold/italic font formatting and/or fill background. Font color in CF rules is limited by NPOI 2.7.3's `IFontFormatting` surface; callers needing full color control reach through `ISheet.Underlying.SheetConditionalFormatting`.
2. **Formula-based** — highlight cells where an arbitrary formula evaluates to TRUE.
3. **Color scale** — 2-color or 3-color gradient. Colors set via XSSF's RGB path.

Multiple rules per `AddConditionalFormatting` call are applied in order; Excel evaluates top-to-bottom. Multiple calls produce separate CF groups.

**Out of scope:** data bars, icon sets, top/bottom N, unique/duplicate, text contains. All reachable via `Underlying`.

### 6.2.10 Charts — I-75

```csharp
public enum ChartType { Line, Bar, Column, Pie, Scatter, Area }

public interface IChart
{
    ISheet Sheet { get; }
    ChartType Type { get; }
    void SetTitle(string title);
    XSSFChart Underlying { get; }
}

// On ISheet:
IChart AddChart(ChartType type, string startCell, string endCell,
    string categoryRange, string valueRange, string? title = null);
```

**I-75 (added 2026-05-26):** Single-series chart facade covering the six chart types NPOI 2.7.3 exposes via `IChartDataFactory`: line, bar, column, pie, scatter, area. The chart is anchored between two cells, with data sourced from a category range (labels/X-axis) and a value range (Y-axis).

**Multi-series charts** are not directly supported — the single-series facade covers the most-common use case. Callers needing multiple series, secondary axes, custom formatting, or chart-type combinations reach through `IChart.Underlying` to access the full `XSSFChart` / `CT_ChartSpace` API.

**Scatter charts** use `FromNumericCellRange` for both axes (X must be numeric); all other chart types use `FromStringCellRange` for categories.

### 6.2.11 Excel-correct defaults + DefaultColumnWidth — I-78

```csharp
// On ISheet:
double? DefaultColumnWidth { get; set; }
```

**I-78 (added 2026-05-27):** Two internal fixes to match Excel's native workbook output, plus a new public property:

1. **Normal cellStyle entry.** NPOI 2.7.3 omits `<cellStyle name="Normal" builtinId="0"/>` from styles.xml. Without it, Excel cannot resolve the Normal style → font index 0 → Maximum Digit Width chain, causing default column widths to display incorrectly. `Workbook.Create()` now ensures this entry exists via `NpoiInternals.GetStylesheet()`.

2. **Suppress `defaultColWidth`.** NPOI writes `defaultColWidth="8.43"` into `<sheetFormatPr>`, but Excel-authored files omit it. When present, Excel uses the literal float value (which rounds differently than the font-derived value), displaying columns at 7.71 instead of 8.43. `AddSheet()` now sets `defaultColWidth = 0` on new sheets, which NPOI serializes as absent.

3. **`ISheet.DefaultColumnWidth`** — nullable double property. `null` (default) means "omit from XML, let Excel derive from font." Non-null writes the attribute explicitly.

### 6.2.12 Row height + two-cell image anchoring — I-76

```csharp
// On IRow:
float HeightInPoints { get; set; }

// On ISheet (new overloads):
IPicture AddPicture(string startCell, string endCell, byte[] data, ImageFormat format);
IPicture AddPicture(string startCell, string endCell, byte[] data);
```

**I-76 (added 2026-05-27):** Two gaps closed for form-layout reproduction:

1. **`IRow.HeightInPoints`** — get/set row height in points. Delegates to NPOI's `IRow.HeightInPoints`. Enables precise vertical layout control for form-style sheets where different rows have different heights.

2. **Two-cell `AddPicture` overload** — anchors an image between `startCell` (top-left) and `endCell` (bottom-right), stretching to fill the anchor region. Unlike the single-cell overload (I-52) which renders at natural pixel size, this variant gives layout control when the anchor position matters more than pixel fidelity.

### 6.2.14 Read-side introspection: themes + drawings — I-81

```csharp
// On IWorkbook:
byte[]? GetThemeXml();
Color? ResolveThemeColor(int index, double tint = 0);
Color? ResolveThemeColor(ThemeColor color);
Color? ResolveThemeColor(string schemeName, double tint = 0);
int? GetThemeLineWidthEmu(int oneBasedIdx);

// On ISheet:
IReadOnlyList<IPicture>   Pictures   { get; }
IReadOnlyList<IConnector> Connectors { get; }

// On IPicture (added): FromCell, ToCell, Dx1..Dy2, Data
// On IConnector (added): FromCell/ToCell/Dx1..Dy2, FlipH/V, HeadEnd/TailEnd,
//                        LineColor, LineSchemeColor, LineWidthPoints,
//                        LineStyleRefIndex
```

**I-81 (added 2026-05-29):** The library was strong on the write side and
thin on the read side. Every consumer that needed to introspect an
existing workbook (validators, converters, reproduction tools) was
reaching through `.Underlying` for the same handful of things —
the working-agreement signal that they belong in the library. I-81 lands
the symmetric read surface:

1. **`GetThemeXml()`** — counterpart to I-79's `SetThemeXml`. Reads the
   theme part directly from the OPC package by name, not via NPOI's
   `XSSFWorkbook.GetTheme()` (which returns a value cached at construction
   and doesn't reflect a post-Open `SetThemeXml`).

2. **`ResolveThemeColor`** — three overloads cover the three callers
   already in the wild: an integer index (OOXML cell-color encoding:
   `0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6, 10=hlink, 11=folHlink`,
   matching `ThemeColor.Index`), a `ThemeColor`, and a drawing scheme name
   (`"dk1"`/`"accent3"`/`"tx1"`, with `tx1`/`bg1`/`tx2`/`bg2` aliases for
   `dk1`/`lt1`/`dk2`/`lt2`). `tint` is applied with Excel's actual tint
   algorithm (HLS lightness with the `tint<0 ? L*(1+tint) : L*(1-tint)+t`
   formula from the OOXML spec) — **not** NPOI's `RGBWithTint`, which
   disagrees with Excel.

3. **`GetThemeLineWidthEmu(int)`** — for connectors and shapes whose width
   comes from the theme's `lnStyleLst` via a `style/lnRef/@idx` reference.

4. **`ISheet.Pictures` / `Connectors`** — drawing-order read-only lists,
   wrapping existing `XSSFPicture`/`XSSFConnector` into the existing
   `IPicture`/`IConnector` facades via `FromExisting` factories that
   derive the missing metadata (image format from MIME; connector type
   from preset geometry).

5. **`IPicture`/`IConnector` property additions** complete the round-trip:
   anchor cells (A1, mirroring the write-side `AddPicture`/`AddConnector`
   string params), EMU offsets, image bytes for pictures, and for
   connectors the flip flags, arrowhead types, explicit line color/width,
   and (crucially) the `lnRef` scheme color name + index so callers can
   resolve theme-styled lines via the new workbook helpers above.

Internally, the workbook caches a parsed `ThemeInfo` lazily on first
`ResolveThemeColor` / `GetThemeLineWidthEmu` access; `SetThemeXml`
invalidates the cache.

### 6.2.15 Engine swap: NPOI → Open XML SDK — I-82

```csharp
// New, additive during the swap (decision I-82):
public static IWorkbook Workbook.CreateOoxml(WorkbookOptions? options = null);
public static IWorkbook Workbook.OpenOoxml(string path, WorkbookOptions? options = null);
public static IWorkbook Workbook.OpenOoxml(Stream stream, bool leaveOpen = true, WorkbookOptions? options = null);

// On IWorkbook (additive):
DocumentFormat.OpenXml.Packaging.SpreadsheetDocument? OpenXmlDocument { get; }
```

**I-82 (added 2026-05-31):** NetXlsx swaps its engine from NPOI 2.7.3 to
Microsoft's **Open XML SDK** (`DocumentFormat.OpenXml`, pinned 3.5.1). This is
the headline change for **v2.0.0**.

*Why.* The fidelity-patching pattern of recent slices is a symptom of NPOI's
lossy read surface (`FillPattern` throws on null `patternType`, `XSSFColor.RGB`
is null for indexed/theme refs, `NumFormattingRuns` counts unformatted runs with
null fonts, `RGBWithTint` disagrees with Excel). The data was always in the
OOXML; NPOI obscured it. Open XML SDK is schema-complete (generated from the
OOXML XSDs), .NET-native, MIT-licensed, AOT/trim-friendly, and Microsoft-
maintained. It also retires the OSMF license ceiling at NPOI 2.8+. Chosen over
binding ClosedXML (the EV-ranked option in `docs/long-term.md`) so NetXlsx owns
its full surface end-to-end rather than becoming a wrapper-of-a-wrapper — matching
the "no escape-hatch surprises" philosophy.

*Strategy — parallel engine, late cutover.* The swap does **not** rewrite the
default engine in place. Instead the SDK engine grows additively, behind new
factory methods, alongside the untouched NPOI engine:

- `IWorkbook.Underlying` keeps returning NPOI's `XSSFWorkbook` for the **entire**
  swap. On the SDK engine it throws `NotSupportedException`; the SDK escape hatch
  is the new `IWorkbook.OpenXmlDocument` (`null` on the legacy engine, non-null on
  the SDK engine — this is how callers discriminate engines).
- `Workbook.Create()` / `Open()` stay NPOI-backed; `Workbook.CreateOoxml()` /
  `OpenOoxml(...)` reach the SDK engine.
- New conformance tests live in `tests/NetXlsx.OoxmlEngine.Tests` so the full
  `NetXlsx.Tests` suite stays green every session. `main` is **never** left
  non-compiling — that invariant is the whole point of the parallel approach.

*The breaking change is deferred to a single, focused cutover slice.* At v2.0.0
cutover: `IWorkbook.Underlying`'s return type changes to `SpreadsheetDocument`,
`Create()`/`Open()` route to the SDK engine, the ~34 NPOI reach-through tests in
`NetXlsx.Tests` migrate in one batch, the `NetXlsx.OoxmlEngine.Tests` conformance
suite folds into the main suite, and the NPOI `PackageReference` + its transitive
security overrides are dropped. The cutover is gated on the **full v1.3 behavioral
suite passing against the SDK engine** — no exceptions.

*Slice order* (each ends green on its tests before the next starts): ✅ foundation
(create/open/save/dispose, AddSheet, enumeration) → ✅ cells & rows (string/number/
bool values, indexers, Range, AppendRow/Row, Column) → ✅ cell styles (new
`OoxmlStylePool` re-targets OOXML schema types — `font`/`fill`/`border`/`numFmt`/
`xf` — with `CellStyle`→`cellXfs` dedup per #4; wires `ICell.Style`/`NumberFormat`/
`GetStyle`/`ApplyNamedStyle`, `IRange.Apply`, `IColumn` width/hidden/style, `IRow`
height/hidden; unblocks `SetDate`/`SetTime`/`SetDuration` + `GetDate`/`Kind==Date`
via the 1900/1904 serial + date-format detection; applies `DefaultFont`/`Excel1904`
on create) → ✅ rich text (`SetRichText`/`GetRichText` as inline `<is>` runs; each
run's `<rPr>` built by the style pool's run-property helper; an empty-style run
gets no `<rPr>` so it inherits the cell font — lesson #10 — and the OPC-preservation
gate, lesson #13, is confirmed on the SDK engine via `OpcPreservationTests`) → ✅
schema-validation gate (cross-cutting conformance — `OpenXmlValidator` over engine
output, target `Microsoft365`; found the engine already schema-clean, including the
`<rPr>` child order the handoff flagged) → 🟡 structure (split across sessions):
✅ merges + named ranges done; **panes / visibility / tab color / protection /
grouping** ←NEXT → drawings →
CF/validation/tables/autofilter/sort → charts → streaming (the `OpenXmlWriter`
forward-only shape may need small public-API tweaks — surface those as their own
decision rows) → source-gen runtime helpers.

*Cell-styles slice — deferred within the surface* (tracked here so a later slice
picks them up, not lost): (a) `GetString` on a date cell returns the raw serial,
not a format-rendered string — the SDK engine has no number-format renderer yet
(the NPOI engine uses `DataFormatter`); the cutover gate requires this. (b) Named
styles resolve in-memory (`RegisterStyle`/`ApplyNamedStyle` work within a session)
but do not persist into OOXML's `cellStyles` panel across save/open — the SDK
equivalent of the NPOI engine's I-67 round-trip. (c) `IColumn.AutoSize` needs
font-metric measurement, which the SDK engine does not carry; it lands with the
font-metrics work, not the styles slice.

*Schema-validation gate — I-82 sub-decision (added 2026-05-31).* The engine
swap's founding premise is that the Open XML SDK is schema-complete, so engine
output is correct "for free." That premise is now *enforced*, not assumed: a
reusable gate (`OpenXmlValidationGate.AssertValid` in
`tests/NetXlsx.OoxmlEngine.Tests`) runs `DocumentFormat.OpenXml.Validation.
OpenXmlValidator` over a workbook's live `OpenXmlDocument` and asserts zero
errors, dumping each error's part URI / element XPath / id / description on
failure. Fixtures cover every landed feature (scalar values; fonts/fills/borders/
numFmts/alignment; 1900- and 1904-epoch dates/times/durations; rich text including
the empty-style inheriting run) plus an `OpenOoxml`→`Save` round-trip. Future
slices add their fixtures to the gate, so it widens with the engine.

- **Validation target: `FileFormatVersions.Microsoft365`.** Microsoft365 is the
  most current target and matches the engine's modern-Excel orientation (it
  round-trips x14/x15 extension parts unmodeled). All created-workbook fixtures
  also validate clean under the conservative `Office2019` alternative; Microsoft365
  is the standing gate. This is a conformance-policy decision and adds no public
  symbol (no `PublicAPI.Unshipped` entry).
- **The rich-text `<rPr>` child-order suspicion is disproven by the schema
  oracle.** The handoff carried `<rPr>` child order as the prime suspect (the
  slice emits `b/i/u/sz/color/rFont`, copied from the order-*independent* styles
  `<font>`). The validator settles it: it does **not** constrain `CT_RPrElt` child
  order — the current order, the strict-ECMA order, and a deliberately scrambled
  order all validate clean — whereas it **does** constrain `CT_Font` order in
  `styles.xml` (a scrambled `<font>` raises `Sch_UnexpectedElementContentExpecting
  Complex`). The engine's `<rPr>` emit order is therefore schema-valid as written,
  and the font path's order was already correct. No reorder was needed; nothing was
  changed in the engine for this slice — the gate found the engine already clean.
- **Known validation finding (real stress files, not engine output):** of the five
  real stress files, four round-trip schema-clean; `ANIMAL_STRAW_HOLDERS_PSS`
  carries one error — `x14:workbookPr/@defaultImageDpi='32767'` in `workbook.xml`,
  which exceeds the attribute's enumeration constraint. That value is **authored in
  the source file** (by whatever produced the PSS workbook) and the engine
  OPC-preserves it byte-for-byte per lesson #13. It is not engine-generated;
  "correcting" it would violate the preservation contract. The committed gate uses
  a synthetic round-trip (the project commits no binary fixtures — decision I18
  option b — and the stress files live only in the operator's Downloads); the
  stress-file validation is operator-side evidence.

*Structure slice — merges + named ranges (I-82 sub-slice, added 2026-05-31).* The
first half of the structural surface lands on the SDK engine; it adds no public
symbol (every member is an existing interface member newly implemented internally,
so no `PublicAPI.Unshipped` entry and the snapshot gate stays green). The slice is
split across sessions per its size — panes / visibility / tab color / protection /
grouping are the next session's half.

- **Merges** (`OoxmlSheet.Merges.cs`): `<mergeCells>` is inserted *after*
  `<sheetData>` in `CT_Worksheet`'s strict sequence (SDK-quirk #3), mirroring how
  `<cols>` is inserted *before* it. `CT_MergeCells` requires ≥1 child, so unmerging
  the last region drops the container rather than leaving a schema-invalid childless
  `<mergeCells/>`; the optional `@count` is kept in sync. Contract matches the NPOI
  engine exactly: 1×1 merge is a no-op (I-38), an overlapping merge throws
  `InvalidOperationException` (§6.4), `UnmergeCells` of a non-exact range is a silent
  no-op, `MergedRanges` returns canonical `A1:C3` strings, and `MergeCellsStyled`
  styles every cell in the range before merging (lesson #4 — merged-region borders
  render from the boundary cells, so cells emit before merges).
- **Named ranges** (`OoxmlWorkbook.Names.cs` + `OoxmlNamedRange.cs`):
  `<definedNames>` is inserted between `<sheets>` and `<calcPr>` in `CT_Workbook`.
  A `localSheetId` (0-based document-order sheet index) carries sheet scope and is
  resolved back to the sheet name on read. The leading `=` is stripped; names must
  be unique workbook-wide case-insensitively regardless of scope (the NPOI engine's
  documented constraint — its duplicate message contains both "already exists" and
  "unique workbook-wide" so the existing `NamedRangeApiTests` pass at cutover). The
  SDK has no built-in name validator, so the documented Excel name rules (start with
  letter/underscore; letters/digits/underscore/period; not a cell reference like A1)
  are enforced in `ValidateDefinedName` — reproducing what the NPOI engine delegated
  to `XSSFName.ValidateName`.
- Both fixture families are added to the schema-validation gate; the engine output
  validates clean under `Microsoft365` (no engine quirk surfaced). All NPOI-engine
  `FreezeMergeHiddenTests` (merge half) and `NamedRangeApiTests` were cross-checked
  against this SDK behavior — every assertion is satisfied, so this surface's cutover
  is already de-risked.

### 6.2.13 Connectors — I-79, I-80

```csharp
public enum ConnectorType { Straight = 96, Bent = 98, Curved = 102 }
public enum ConnectorEnd { None, Triangle, Stealth, Diamond, Oval, Arrow }

public interface IConnector
{
    ISheet Sheet { get; }
    ConnectorType Type { get; }
    XSSFConnector Underlying { get; }
}

// On ISheet:
IConnector AddConnector(ConnectorType type, string startCell, string endCell,
    Color? lineColor = null,
    int dx1 = 0, int dy1 = 0, int dx2 = 0, int dy2 = 0,
    bool flipH = false, bool flipV = false,
    ConnectorEnd headEnd = ConnectorEnd.None, ConnectorEnd tailEnd = ConnectorEnd.None,
    double? lineWidthPoints = null);
```

**I-79 (added 2026-05-27):** Initial connector support — `AddConnector(type, startCell, endCell, lineColor)` returning the raw `XSSFConnector`.

**I-80 (added 2026-05-27):** Connector support reworked for faithful reproduction. Three problems with the I-79 shape:

1. **Wrong geometry.** `ConnectorType` values were old POI `ShapeTypes` constants (20/32/38), but `XSSFConnector.ShapeType` takes an `ST_ShapeType` ordinal. Value 20 is `star8`, not a connector — every I-79 connector rendered as an 8-pointed star. Values are now the correct `ST_ShapeType` ordinals (`straightConnector1 = 96`, `bentConnector3 = 98`, `curvedConnector3 = 102`).

2. **No EMU offsets.** I-79 hard-coded `(0,0,0,0)` anchor offsets, so a connector confined to one cell (the common case for a short arrow over a merged region) collapsed to zero length and was invisible. `AddConnector` now accepts `dx1..dy2` like the two-cell `AddPicture` overload.

3. **No arrowheads / flip / width.** A "Straight **Arrow** Connector" needs a `tailEnd`/`headEnd` decoration; direction is encoded as `flipH`/`flipV` in the transform; weight comes from the theme line-style ref (e.g. `lnRef idx="2"` ≈ 2 pt). I-80 surfaces `headEnd`/`tailEnd` (`ConnectorEnd`), `flipH`/`flipV`, and `lineWidthPoints` so the generated code stays pure NetXlsx — no reaching through `.Underlying`.

The return type is now the `IConnector` facade (consistent with `IShape`/`IPicture`/`IChart`); the raw `XSSFConnector` remains reachable via `.Underlying`. Theme-based line *color* is left to the caller — `lineColor` sets an explicit `solidFill` which Excel renders over the style `lnRef` color.

### 6.2.12 Drawings / shapes — I-74

```csharp
public enum ShapeType
{
    Rectangle = 5, RoundedRectangle = 26, Ellipse = 35,
    Line = 1, Triangle = 3, Diamond = 6,
}

public interface IShape
{
    ISheet Sheet { get; }
    ShapeType Type { get; }
    XSSFSimpleShape Underlying { get; }
}

// On ISheet:
IShape AddShape(ShapeType type, string startCell, string endCell,
    Color? fillColor = null, Color? lineColor = null);
```

**I-74 (added 2026-05-26):** Minimal shape facade — covers the six most common shapes (rectangle, rounded rectangle, ellipse, line, triangle, diamond). Anchored between two cells. Fill and line colors are optional; no-fill is the default. Advanced shape properties (rotation, line width, gradient, text, etc.) reach through `IShape.Underlying`.

Exotic shapes (arrows, callouts, stars, connectors, freeforms) are accessible via `ISheet.Underlying.CreateDrawingPatriarch()` which returns the `XSSFDrawing`.

### 6.2.11 Sorting helpers — I-72

```csharp
public sealed class SortKey
{
    public int Column { get; }       // 1-based
    public bool Ascending { get; }
    public static SortKey Asc(int column);
    public static SortKey Desc(int column);
}

// On ISheet:
void SortRange(string a1Range, params SortKey[] keys);
```

**I-72 (added 2026-05-26):** Pure facade — physically reorders cell values and styles within a range by the specified sort keys. No OOXML `sortState` metadata is written; the sort is immediate and in-memory.

Sort order matches Excel: numbers before strings, blanks last (ascending), case-insensitive string comparison. Multiple keys are applied in declared order (primary, then secondary, …). The range's first row is NOT treated as a header — include only data rows.

Implementation snapshots all cell values/styles in the range, sorts the snapshot array, and writes sorted values back. Cell styles travel with their values.

### 6.2.10 Split panes — I-70

```csharp
// On ISheet:
void CreateSplitPane(int xSplitTwips, int ySplitTwips);
```

**I-70 (added 2026-05-26):** Complements the existing `FreezePane` (decision §6.4) with a draggable split. Unlike freeze panes, split panes allow the user to resize the split interactively in Excel. Parameters are in twips (1/20th of a point), matching NPOI's `CreateSplitPane` and the OOXML `<pane>` element's coordinate system.

`CreateSplitPane` replaces any prior freeze or split on the sheet. The active pane defaults to `LowerRight`. Callers needing a specific active pane or the leftmostColumn / topRow hints reach through `ISheet.Underlying.CreateSplitPane(x, y, leftCol, topRow, activePane)`.

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

### 6.4.1 Tables (ListObject) — I-51

```csharp
public interface ITable
{
    string Name { get; }                 // codename (Excel name rules)
    string DisplayName { get; set; }
    string Address { get; }              // "A1:D10"
    ISheet Sheet { get; }
    IReadOnlyList<string> ColumnNames { get; }
    bool HasTotalsRow { get; }
    string? StyleName { get; set; }
    NPOI.XSSF.UserModel.XSSFTable Underlying { get; }
}

// On ISheet:
ITable AddTable(string a1Range, string name, string? style = null);
IReadOnlyList<ITable> Tables { get; }
bool TryGetTable(string name, [MaybeNullWhen(false)] out ITable table);

public static class TableStyles
{
    public const string Light1 = "TableStyleLight1";
    public const string Medium2 = "TableStyleMedium2";   // Excel default
    public const string Dark9 = "TableStyleDark9";
    // ... curated subset of common names ...
}
```

**I-51 (added 2026-05-22):** Excel Tables are sheet-scoped structured ranges with a header row, optional totals row, optional style, and OOXML-mandatory AutoFilter. `ISheet.AddTable` requires:

- A range with **at least two rows** (one header row, one data row).
- A **non-empty string** in every header cell — these become the table's column names. No "auto Column1/Column2/…" path; reject explicitly.
- A **valid Excel name** (letters / digits / underscores; must start with letter or underscore; cannot collide with an A1 reference) **unique workbook-wide** (case-insensitive). Table names share the namespace with named ranges and other tables.

Style is a `string?` keyed against Excel's built-in style names (e.g., `TableStyleMedium2`). A `TableStyles` static class provides constants for the common subset; arbitrary Excel-recognized strings work. `null` (the default) applies no style.

`HasTotalsRow` is **read-only** in v1.1 — adding totals requires per-column totals-row functions which is a larger surface; defer to v1.2 or reach through `.Underlying`. `RemoveTable` is also deferred: NPOI 2.7.3's `XSSFSheet` has no `RemoveTable` method, so removal requires package-part manipulation. Defer to v1.2 or NPOI 3.x.

The implementation populates `CT_Table.tableColumns` directly rather than calling `XSSFTable.CreateColumn`, because NPOI 2.7.3's `CreateColumn` throws when the underlying `tableColumn` list is uninitialized (NPOI surprise; captured in `implementation-notes.md`).

### 6.4.2 Image embedding — I-52

```csharp
public enum ImageFormat { Png, Jpeg }

public interface IPicture
{
    ISheet Sheet { get; }
    ImageFormat Format { get; }
    NPOI.XSSF.UserModel.XSSFPicture Underlying { get; }
}

// On ISheet:
IPicture AddPicture(string a1Cell, byte[] data, ImageFormat format);
IPicture AddPicture(string a1Cell, byte[] data);   // magic-byte detect

public sealed class UnsupportedImageFormatException : WorkbookException { ... }
```

**I-52 (added 2026-05-22):** v1.1 image embedding is deliberately narrow:

- **Two formats: PNG and JPEG.** The two formats Excel reads on every platform without theme-color quirks. Other formats (GIF, BMP, TIFF, EMF, WMF) reach through `IWorkbook.Underlying.AddPicture` + `ISheet.Underlying.CreateDrawingPatriarch`.
- **Single-cell anchor.** The image's top-left corner sits at the supplied A1 cell; the picture is rendered at its **natural pixel size** via NPOI's `XSSFPicture.Resize()`. Multi-cell anchoring, pixel offsets, alt-text, and rotation reach through `IPicture.Underlying`.
- **Auto-detect from magic bytes.** The 2-arg overload reads the first bytes (`89 50 4E 47 ...` for PNG, `FF D8 FF ...` for JPEG) and dispatches. Unknown formats throw `UnsupportedImageFormatException`. The 3-arg overload accepts an explicit `ImageFormat` and skips detection.

`Resize()` is essential — without it, NPOI anchors the picture to a single-cell extent, stretching/shrinking the image to fit the (typically small) cell. Calling `Resize()` makes the picture's to-cell match the image's pixel dimensions, producing the expected display.

### 6.4.3 Sheet protection — I-53

```csharp
public sealed record SheetProtection
{
    public bool LockFormatCells { get; init; }
    public bool LockFormatColumns { get; init; }
    public bool LockFormatRows { get; init; }
    public bool LockInsertColumns { get; init; }
    public bool LockInsertRows { get; init; }
    public bool LockInsertHyperlinks { get; init; }
    public bool LockDeleteColumns { get; init; }
    public bool LockDeleteRows { get; init; }
    public bool LockSelectLockedCells { get; init; }
    public bool LockSelectUnlockedCells { get; init; }
    public bool LockSort { get; init; }
    public bool LockAutoFilter { get; init; }
    public bool LockPivotTables { get; init; }
    public bool LockObjects { get; init; }
    public bool LockScenarios { get; init; }
    public static SheetProtection Default { get; }   // all false
    public static SheetProtection LockAll { get; }   // all true
}

// On ISheet:
void Protect(string? password = null, SheetProtection? options = null);
void Unprotect();
bool IsProtected { get; }
```

**I-53 (added 2026-05-22):** Sheet protection is a **UX guard, not security**. Excel's sheet-protection password is hashed with a weak algorithm widely known to be brute-forceable. Use for "stop accidental edits", not "stop a determined attacker." Documented in the type's XML doc.

`Protect()` with no arguments enables the protection flag without a password (Excel will block edits but accept any "unprotect" request immediately). `Protect(password: "...")` adds the (weak) password hash. Both forms accept an optional `SheetProtection` record with 15 granular `Lock*` flags mirroring NPOI 2.7.3's `XSSFSheet.Lock*(bool)` methods.

**NPOI surprise:** `XSSFSheet.ProtectSheet(null)` is NPOI's *unprotect* operation, not "protect without password". The no-password path therefore manipulates `CT_SheetProtection` directly (`sp.sheet = true; sp.scenarios = true; sp.objects = true;` — matching the side effects of `ProtectSheet(non-null)`). Captured in `implementation-notes.md`.

### 6.4.4 Data validation — I-55

```csharp
public sealed class DataValidation
{
    // List (dropdown)
    public static DataValidation List(params string[] values);
    public static DataValidation ListFromRange(string formula);
    // Integer / decimal
    public static DataValidation IntegerBetween(int min, int max);
    public static DataValidation IntegerEqual(int value);
    public static DataValidation IntegerGreaterThan(int value);
    public static DataValidation IntegerLessThan(int value);
    public static DataValidation DecimalBetween(double min, double max);
    // Date / text / custom
    public static DataValidation DateBetween(DateOnly start, DateOnly end);
    public static DataValidation TextLengthAtMost(int max);
    public static DataValidation TextLengthAtLeast(int min);
    public static DataValidation Custom(string formula);
}

// On ISheet:
void AddValidation(string a1Range, DataValidation validation);
```

**I-55 (added 2026-05-22):** Data validation is exposed via a single sealed class with static factories — no public constructor, no exposed inheritance hierarchy. Each factory captures the NPOI helper-method call site as a lambda; the captured rule materializes against the sheet's `IDataValidationHelper` at `AddValidation` time.

**Date validation uses the `DATE(yyyy,m,d)` formula form** rather than a locale-specific literal. Excel evaluates `DATE(...)` deterministically across all locales; a literal like `"5/22/2026"` would be misparsed on machines with European date formats. Captured in `implementation-notes.md`.

**Out of v1.1 scope:** time-of-day validation, "not between" / "not equal" operators, formula-driven decimal/integer constraints, error-style customization (Stop / Warning / Information), per-validation prompt + error messages. These reach through `ISheet.Underlying.GetDataValidationHelper()`.

### 6.4.5 AutoFilter — I-56

```csharp
// On ISheet:
void SetAutoFilter(string a1Range);
void ClearAutoFilter();
bool HasAutoFilter { get; }
string? AutoFilterRange { get; }
```

**I-56 (added 2026-05-22):** Standalone AutoFilter for ranges that are not tables. `SetAutoFilter` applies Excel's dropdown filter over a range (the first row of the range becomes the header); replacing an existing AutoFilter just updates the range. `ClearAutoFilter` removes it.

**Out of v1.1 scope:** per-column filter criteria (filter A="X" AND B>0, etc.). Excel's filter criteria model is rich (text equals/contains, top-N, color, date range, custom expression) — exposing it as a v1.1 surface would be a significant chunk of API. Callers reach through `ISheet.Underlying.GetCTWorksheet().autoFilter.filterColumn` for now.

**NPOI surprise:** `CT_Worksheet` in NPOI 2.7.3 exposes the `autoFilter` element as a direct property, not via `IsSetX` / `UnsetX` accessors. `ClearAutoFilter` therefore assigns `null` to the property to remove the element from the serialized XML. The auxiliary `_FilterDatabase` built-in name (created by NPOI's `SetAutoFilter`) is left in place when clearing — Excel tolerates a stale name pointing at an absent autoFilter, and pruning it would require walking the workbook's name table.

### 6.4.7 Table totals row — I-64 (v1.2)

```csharp
public enum TotalsRowFunction
{
    None, Sum, Min, Max, Average, Count, CountNumbers, StdDev, Var, Custom,
}

// On ITable:
void AddTotalsRow();
void RemoveTotalsRow();
void SetColumnTotal(string columnName, TotalsRowFunction function);
void SetColumnTotal(string columnName, string customFormula);
void SetColumnTotalLabel(string columnName, string label);
```

**I-64 (added 2026-05-22):** v1.1 slice 2 (decision I-51) made `ITable.HasTotalsRow` read-only because adding totals requires per-column functions and the full surface didn't fit in slice scope. v1.2 closes the gap.

`AddTotalsRow` extends the table's `@ref` by one row downward and sets `CT_Table.totalsRowCount = 1`. When the table has an AutoFilter (auto-applied by NPOI's `SetCellReferences`), the autoFilter range is trimmed back to exclude the totals row — matching Excel's default behavior where filters skip the totals.

`SetColumnTotal` writes both the OOXML metadata (`CT_TableColumn.totalsRowFunction`) **and** the actual cell formula (`SUBTOTAL(code, TableName[ColumnName])`). The dual write means the total renders correctly in any conforming viewer, not just one that auto-populates from the function metadata on open. SUBTOTAL is invoked in its **100-series form** (`101..110`) — the variant that skips AutoFilter-hidden rows, matching Excel's table-totals behavior.

The structured-reference form (`TableName[ColumnName]`) is used in the formula rather than an absolute cell range so the totals auto-update when rows are added to the table.

**Out of scope (deferred to a later slice if needed):** column names containing structured-reference special characters (`#`, `[`, `]`, `'`, `@`, whitespace) would need quoting in the formula body. v1.2 covers the common-case unquoted form; names that need quoting reach through `Underlying`.

`SetColumnTotalLabel` writes a text label to the cell (typically `"Total"` on the leading column) and explicitly clears any function metadata — the label takes precedence in Excel's rendering.

`RemoveTotalsRow` is the inverse: clears per-column functions and labels, blanks the totals-row cells in the table's column range, shrinks the table's `@ref` by one row.

### 6.4.6 Table removal — I-63 (v1.2)

```csharp
// On ISheet:
void RemoveTable(ITable table);
```

**I-63 (added 2026-05-22):** v1.1 slice 2 (decision I-51) deferred `RemoveTable` because NPOI 2.7.3's `XSSFSheet` did not expose a `RemoveTable` method (the upstream source has one; the 2.7.x binary line never published it). v1.2 closes the gap by performing the three-step removal directly:

1. Drop the matching `<tablePart>` entry from `CT_Worksheet.tableParts` (matched by the table's package-relationship id).
2. Remove the package relationship + part via `POIXMLDocumentPart.RemoveRelation`.
3. Drop the cached entry from `XSSFSheet`'s internal `tables` dictionary so subsequent `GetTables()` snapshots don't surface the removed entry.

Steps 2 and 3 require crossing NPOI's protection boundary — `RemoveRelation` is `protected` and the `tables` field is `private`. Both crossings are centralized in `src/NetXlsx/Internal/NpoiInternals.cs`, with `MethodInfo` / `FieldInfo` cached as `static readonly` fields so each reflection lookup happens once. A future NPOI 3.x bump that exposes `RemoveTable` publicly removes both reflection calls; until then, this is the narrowest workable surface.

**Validation:** `XssfSheet.RemoveTable` rejects foreign-table handles (the table's relationship id won't resolve on a different sheet) and rejects already-removed handles for the same reason — both throw `ArgumentException`. A second `RemoveTable(t)` on a freshly-removed handle is not idempotent-silent; it surfaces the stale handle loudly. Calling code that wants idempotency should check `sh.Tables.Contains(t)` first.

### 6.4.8 Per-column AutoFilter criteria — I-66 (v1.2)

```csharp
public sealed class FilterCriteria
{
    // Operators (single condition)
    public static FilterCriteria EqualTo(string value);
    public static FilterCriteria NotEqualTo(string value);
    public static FilterCriteria GreaterThan(double value);
    public static FilterCriteria GreaterThanOrEqual(double value);
    public static FilterCriteria LessThan(double value);
    public static FilterCriteria LessThanOrEqual(double value);

    // String pattern (encoded as Equal with wildcards)
    public static FilterCriteria Contains(string substring);
    public static FilterCriteria DoesNotContain(string substring);
    public static FilterCriteria StartsWith(string prefix);
    public static FilterCriteria EndsWith(string suffix);

    // Numeric range — two-condition AND
    public static FilterCriteria Between(double min, double max);

    // Combinators (Excel limits to 2 conditions per column)
    public FilterCriteria And(FilterCriteria other);
    public FilterCriteria Or(FilterCriteria other);
}

// On ISheet:
void SetAutoFilterColumn(int columnOffset, FilterCriteria criteria);
void ClearAutoFilterColumn(int columnOffset);
```

**I-66 (added 2026-05-22):** v1.1 slice 7 (decision I-56) shipped range-only AutoFilter — `SetAutoFilter(range)` shows the dropdown arrows but no per-column criteria. v1.2 closes the gap for the **custom-filter** OOXML variant.

**Scope (v1.2):** custom-filter conditions only — operator + value pairs joined by AND or OR. Covers equality (eq, ne), ordering (gt, ge, lt, le), and string patterns (contains / startsWith / endsWith via Excel's `*` / `?` wildcards). The `Between` factory composes two conditions with AND. Excel limits filter columns to **two conditions max**; chaining a third `And` / `Or` throws `InvalidOperationException`.

**Deferred:**
- **Explicit-value list filter** (`filters` element, `In(...)` factory). NPOI 2.7.3's `CT_FilterColumn` does not surface the `filters` property — only `customFilters`. Implementing would require XML-node-level reflection. Candidate for v1.3.
- **Top-N filter** (`top10` element). Same NPOI 2.7.3 surfacing limitation.
- **Date-group filter, dynamic filter (relative dates), color filter.** Niche; not in v1.2 scope.

Implementation builds `CT_CustomFilters.customFilter` directly with the operator enum + value. String-pattern factories prefix/suffix the wildcard `*` to encode the pattern; the input string's literal `*` / `?` / `~` are escaped by prefixing with `~` (Excel's filter-language escape).

Column offset is **0-based within the AutoFilter range** — column 0 is the first column of the range, matching OOXML's `colId`. Negative or out-of-range offsets throw `ArgumentOutOfRangeException` with a friendlier message than NPOI would produce. `SetAutoFilterColumn` requires a prior `SetAutoFilter` call (`InvalidOperationException` otherwise).

### 6.4.8.1 `FilterCriteria.In(...)` — I-68 (v1.3)

```csharp
// On FilterCriteria:
public static FilterCriteria In(params string[] values);
public static FilterCriteria In(IEnumerable<string> values);
```

**I-68 (added 2026-05-22):** The v1.2 AutoFilter slice (I-66) deferred the explicit-value-list filter (`<filters>` element) because NPOI 2.7.3's `CT_FilterColumn` proxy does not model that element — only `customFilters`. v1.3 ships `In(...)` with **two-value support** via the existing `customFilters` infrastructure (the same path Excel uses internally for short value lists) and throws `NotSupportedException` for 3+ values.

**Behavior table:**

| Value count | Implementation | Outcome |
|---|---|---|
| 0 (empty) | rejected | `ArgumentException` |
| 1 | reduces to `EqualTo(v[0])` | single `customFilter` |
| 2 | composes `EqualTo(a).Or(EqualTo(b))` | two OR-joined `customFilter` entries |
| 3+ | rejected | `NotSupportedException` naming NPOI 2.7.3 + the `<filters>` element as the cause |

The two-value case is the most-common real-world use ("filter to two regions / quarters / statuses"). Adding the surface now means call sites are ready when an NPOI 3.x bump (or a future XML-emission workaround) lifts the 3+ value limit — no caller-code change required.

**Workaround for 3+ values until then:** reach through `ISheet.Underlying.GetCTWorksheet().autoFilter` and write the OOXML directly, or stage the data so the filter applies to fewer distinct values.

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

### 6.9.1 Custom type converters — I-58

```csharp
public interface ICellConverter<T>
{
    void Write(ICell cell, T value);
    T Read(ICell cell);
}

// On ColumnAttribute:
[Column("Tags", ConverterType = typeof(TagsConverter))]
public List<string> Tags { get; init; }
```

**I-58 (added 2026-05-22):** Lets a `[Worksheet]`-mapped property carry any type the user can write a converter for, escaping the generator's built-in scalar set (string / numeric / bool / DateTime / DateOnly / TimeOnly / TimeSpan).

A configured `ConverterType` **overrides** the built-in `IsSupportedPropertyType` check — the property is treated as emissible regardless of declared type. The generator emits one `static readonly` field per converter property, sharing one instance across all `AddRow` / `ReadRows` calls. The converter's `Write` and `Read` methods are invoked through that cached instance.

**Constraints (compile-time):**
- Converter type must be non-abstract and have a public parameterless constructor.
- Converter type must implement `ICellConverter<T>` where `T` exactly matches the property type.

Violations surface at C# compile time via the cast in the generated field-initializer (`new SomeConverter()` assigned to `ICellConverter<T>` — mismatched `T` → CS0029 / similar). v1.1 does not add a dedicated `NXLS0007` diagnostic for this — the C# error is already clear.

**Why not a workbook-level registry?** Considered (`workbook.Converters.Register<T, TConv>()`); rejected because the source generator runs at compile time and can't see runtime registrations. Per-property attribute keeps the converter visible at the call site and binds at code-emit time.

### 6.13.1 Fuzz harness for the open path — I-60

A standalone test project `tests/NetXlsx.Fuzz/` (xUnit, opt-in via the
`Fuzz` trait) feeds malformed inputs to `Workbook.Open` and asserts:

1. Every input terminates in **bounded time** (2-second per-call cap on
   the bulk sweep). A hang is a finding — resource exhaustion or an
   infinite loop.
2. Every input produces **either** a documented NetXlsx exception
   (`WorkbookException` family + `ResourceLimitExceededException`) **or**
   a documented BCL exception (`InvalidDataException`, `IOException`,
   `FormatException`, `ArgumentException`) **or** an NPOI namespaced
   exception. Anything else is a finding.

Corpus (decision I-60):

- Pure random / all-zero / all-`0xFF` bytes at sizes 0, 1, 16, 256, 4096.
- Empty ZIP, ZIP with random non-OOXML entries, ZIP with truncated
  `[Content_Types].xml`.
- Classic billion-laughs XML expansion bomb in `[Content_Types].xml`.
- High-compression-ratio zip bomb (64 MiB of zeros → ~few KB on disk)
  against a tightened `WorkbookOptions.ReadMaxUncompressedBytes`.
- Bit-flip mutations of a known-good baseline workbook across five
  deterministic seeds.
- 100-iteration bulk random sweep with per-call cancellation timer.

**Driven hardening (post-v1.0):** the initial harness run surfaced
`IndexOutOfRangeException` leaking from NPOI's parsers on truncated /
adversarial input. `Workbook.Open` now translates the runtime-exception
family commonly seen here — `IndexOutOfRangeException`,
`NullReferenceException`, `OverflowException`,
`ArgumentOutOfRangeException` — to `MalformedFileException`. The
underlying NPOI behavior is still arguably a bug there; on our open
path the right user-visible contract is "this file is malformed", not
a leaked runtime exception. Captured in `implementation-notes.md`.

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

### 5.1 Extended benchmark coverage — I-62

The original CI benchmark suite (`benchmarks/NetXlsx.Benchmarks/Benchmarks.cs`) covers the **meso** band — 5K–100K rows, the place where most real workloads live. v1.0 external review flagged two gaps:

1. **Micro** — per-cell timing for scalars, A1-parse, style-pool hit vs miss. Lets a single-cell-cost regression show up immediately rather than being lost in the meso noise.
2. **Macro** — many-sheets (500+) and very-wide layouts (100K cells in a single sheet); plus a streaming variant at 200K rows. Surfaces scaling issues invisible in the meso tier.
3. **Percentile reporting** — P50/P95/P99 emitted to the regression gate alongside mean/median, so a slow-tail regression triggers an alert even when the central tendency is stable.

`benchmarks/NetXlsx.Benchmarks/BenchmarksExtended.cs` ships these as three new `[Config(typeof(CiConfigWithPercentiles))]` classes:

- `MicroBenchmarks` — Cell_SetString, Cell_SetNumber_Double, Cell_SetNumber_Int, Cell_SetBool, Cell_SetDate, Cell_Style_FreshCellStyle (worst-case pool miss), Cell_Style_PoolHit (warm-cache hit), CellAddress_ParseA1, CellAddress_FormatA1.
- `MacroBenchmarks` — Macro_500Sheets, Macro_100kCells_Wide, Macro_Streaming_200kRows.
- `ReadMicroBenchmarks` — Open_OneCell_Read, the read-side floor cost.

All extended benchmarks run under the existing CI regression gate (15% threshold) — new shapes are first-class regression material, not opt-in.

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
