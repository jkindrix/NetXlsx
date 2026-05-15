# NetXlsx — Roadmap

**Status:** Draft for review
**Date:** 2026-05-14

Each capability has a binary answer per release: **Yes** = supported in that version, blank = not in that version, **No** = explicitly out of scope (never).

## Feature matrix

| Capability                              | v1.0 | v1.1 | v2.0 | v3.0 | Never |
|-----------------------------------------|------|------|------|------|-------|
| **Core I/O**                            |      |      |      |      |       |
| Create `.xlsx` workbook                 | Yes  |      |      |      |       |
| Open `.xlsx` workbook                   | Yes  |      |      |      |       |
| Save sync + async-over-sync             | Yes  |      |      |      |       |
| File path API + Stream API              | Yes  |      |      |      |       |
| Streaming write (`IStreamingWorkbook`)  | Yes  |      |      |      |       |
| Streaming read                          |      |      | Yes  |      |       |
| `.xls` (legacy binary)                  |      |      |      |      | No    |
| Encrypted / password-protected files    |      |      |      | Yes  |       |
| **Cells**                               |      |      |      |      |       |
| String/number/date/bool/formula values  | Yes  |      |      |      |       |
| Rich text in cells                      |      | Yes  |      |      |       |
| Hyperlinks                              | Yes  |      |      |      |       |
| Comments (basic)                        | Yes  |      |      |      |       |
| Threaded / modern comments              |      |      | Yes  |      |       |
| **Styling**                             |      |      |      |      |       |
| Font / color / fill / border / align    | Yes  |      |      |      |       |
| Auto style dedup                        | Yes  |      |      |      |       |
| Number format strings                   | Yes  |      |      |      |       |
| Built-in number format constants        | Yes  |      |      |      |       |
| Themes / theme colors                   |      |      | Yes  |      |       |
| Named/reusable styles                   |      | Yes  |      |      |       |
| **Structure**                           |      |      |      |      |       |
| Multiple sheets, rename/reorder/delete  | Yes  |      |      |      |       |
| Merge cells                             | Yes  |      |      |      |       |
| Freeze panes                            | Yes  |      |      |      |       |
| Split panes                             |      |      | Yes  |      |       |
| Hidden rows / columns / sheets          | Yes  |      |      |      |       |
| Auto-size columns                       | Yes  |      |      |      |       |
| Grouping / outlining                    |      |      | Yes  |      |       |
| Named ranges                            | Yes  |      |      |      |       |
| Tables (`ListObject`)                   |      | Yes  |      |      |       |
| **Data features**                       |      |      |      |      |       |
| Formulas (write only — Excel evaluates) | Yes  |      |      |      |       |
| Formula evaluation                      |      |      |      |      | No    |
| Data validation (drop-downs, ranges)    |      | Yes  |      |      |       |
| Conditional formatting                  |      |      | Yes  |      |       |
| AutoFilter                              |      | Yes  |      |      |       |
| Sorting                                 |      |      | Yes  |      |       |
| **Typed mapping**                       |      |      |      |      |       |
| Source-gen `[Worksheet]` writer (ext.)  | Yes  |      |      |      |       |
| Source-gen `[Worksheet]` reader (ext.)  | Yes  |      |      |      |       |
| Custom type converters                  |      | Yes  |      |      |       |
| LINQ provider over sheets               |      |      |      | Yes  |       |
| **Advanced features**                   |      |      |      |      |       |
| Images (PNG/JPEG embed)                 |      | Yes  |      |      |       |
| Charts                                  |      |      | Yes  |      |       |
| Pivot tables (write)                    |      |      |      | Yes  |       |
| Pivot tables (read)                     |      |      |      |      | No    |
| Macros / VBA                            |      |      |      |      | No    |
| Drawings / shapes                       |      |      | Yes  |      |       |
| Sparklines                              |      |      |      | Yes  |       |
| **Output formats**                      |      |      |      |      |       |
| `.xlsx`                                 | Yes  |      |      |      |       |
| `.xlsm` (macro-enabled, passthrough)    |      |      | Yes  |      |       |
| `.xlsb` (binary)                        |      |      |      |      | No    |
| **Protection / security**               |      |      |      |      |       |
| Sheet protection (basic)                |      | Yes  |      |      |       |
| Workbook protection                     |      | Yes  |      |      |       |
| File-level encryption                   |      |      |      | Yes  |       |
| Bounded-resource parsing                | Yes  |      |      |      |       |
| **Developer experience**                |      |      |      |      |       |
| `.Underlying` escape hatch              | Yes  |      |      |      |       |
| `ILogger` integration                   | Yes  |      |      |      |       |
| Public API snapshot tests               | Yes  |      |      |      |       |
| Benchmark suite vs peers                | Yes  |      |      |      |       |
| Golden-file test corpus                 | Yes  |      |      |      |       |
| Sample / cookbook project               | Yes  |      |      |      |       |
| Roslyn analyzers (e.g., "date w/o fmt") |      |      | Yes  |      |       |
| **Platform**                            |      |      |      |      |       |
| `net8.0`, `net9.0`                      | Yes  |      |      |      |       |
| `netstandard2.0` / .NET Framework       |      |      |      |      | No    |
| Native AOT compatible                   | Yes  |      |      |      |       |
| Trim compatible                         | Yes  |      |      |      |       |

## Release themes

- **v1.0 — Foundation.** Core writing and reading, fluent cell API, auto-deduplicating styles, source-generated typed mapping, streaming write, benchmark suite, public-API snapshot tests, cookbook.
- **v1.1 — Common asks.** Tables, rich text, images, sheet/workbook protection, data validation, named styles, custom type converters.
- **v2.0 — Advanced styling & charts.** Conditional formatting, themes, charts, AutoFilter improvements, modern comments, `.xlsm` passthrough, Roslyn analyzers. Breaking changes allowed.
- **v3.0 — Power features.** Pivot tables (write), sparklines, encryption, LINQ provider. Breaking changes allowed.

## Definitions of Done

A release ships when **all** of the following are true:

- [ ] Every "Yes" cell in this release's column is implemented and tested.
- [ ] All performance targets in `design.md §4` are met or improved.
- [ ] Public API snapshot reviewed and approved by maintainers.
- [ ] Golden-file test corpus passes.
- [ ] Manual smoke test on Excel (Windows + Mac), LibreOffice, and Google Sheets passed.
- [ ] Benchmark suite shows no regression > 10% vs prior release.
- [ ] Cookbook updated with new features.
- [ ] `CHANGELOG.md` updated.
- [ ] Migration notes published (if breaking changes).
- [ ] All public APIs have XML documentation.

## Per-release progress tracking

> **Target dates.** All releases are marked `TBD`. Concrete targets will be set after the v1.0 project is scaffolded and we have measured velocity data; this section will be revised in a follow-up commit at that point.

### v1.0 — Foundation (target: TBD)

**Pre-implementation spike** (precondition to locking the design):
- [ ] Style-dedup feasibility benchmark — validate or revise §4 typical/worst-case targets
- [ ] Streaming-write back-pressure measurement on a 1M-row workload
- [ ] Async wrapping cost — measure `Task.Run`-around-NPOI overhead vs sync baseline

**Scaffold** (precondition to writing any library code):
- [ ] Solution + project layout per design §8
- [ ] `Directory.Build.props`: TFMs (`net8.0;net9.0`), nullable enabled, deterministic, signed (S12, S15)
- [ ] `Directory.Packages.props`: central package management, exact NPOI pin (S8)
- [ ] `nuget.config`: public NuGet feed (S10)
- [ ] `.editorconfig` (S5)
- [ ] `LICENSE` (MIT — S9)
- [ ] `CHANGELOG.md` initialized in Keep-a-Changelog format (S21)
- [ ] `CODEOWNERS` (S22)
- [ ] `build/build.ps1` and `build/build.sh` (S18)
- [ ] `.teamcity/` Kotlin DSL pipeline (S17): build, test, golden-file tests, public-API snapshot, benchmarks, pack, publish
- [ ] MinVer wired (S11); first tag `v0.1.0`
- [ ] Strong-name key generated and committed (S12)
- [ ] Source Link configured (S14)
- [ ] PublicApiAnalyzer wired with empty `PublicAPI.Shipped.txt` / `PublicAPI.Unshipped.txt` (S6)
- [ ] xUnit + FluentAssertions wired in test projects (S1, S2)
- [ ] BenchmarkDotNet wired in benchmark project (S3)
- [ ] Empty `docs/npoi-workarounds.md` placeholder (S26)
- [ ] First green CI run on `main`

**Core I/O**
- [ ] `Workbook.Create` returns `IWorkbook`; `Workbook.CreateStreaming` returns `IStreamingWorkbook`
- [ ] `Workbook.Open` / `Workbook.OpenAsync` (path + stream)
- [ ] `Save` / `SaveAsync` (path + stream, `leaveOpen` support)
- [ ] `WorkbookOptions`, `StreamingOptions`

**Cells**
- [ ] String, number (`double`, `decimal`), `DateTime`, `DateOnly`, `TimeOnly`, `TimeSpan`, `bool` setters
- [ ] Formula write (no value cache — §7.8)
- [ ] Typed reads (`GetString`, `GetNumber`, `GetTime`, `GetDuration`, `GetError`, …, `GetValue<T>`)
- [ ] `CellKind` and `CellError` enums + detection
- [ ] Hyperlinks
- [ ] Basic comments

**Styling**
- [ ] Fluent style methods (`Bold`, `Italic`, `Font`, etc.)
- [ ] `CellStyle` value record
- [ ] Internal style cache with value-equality dedup
- [ ] Number formats (string-based) + `NumberFormats` constants

**Structure**
- [ ] `AddSheet` / `RemoveSheet` / `RenameSheet` / `MoveSheet`
- [ ] Sheet indexer by name and index
- [ ] Cell addressing: `["A1"]` and `[r,c]`
- [ ] Range addressing: `Range("A1:C10")` and `Range(r1,c1,r2,c2)`
- [ ] Merge / unmerge
- [ ] Freeze rows / columns / pane
- [ ] Hidden rows / columns / sheets
- [ ] Auto-size columns
- [ ] Named ranges

**Typed mapping**
- [ ] `[Worksheet]` / `[Column]` / `[Ignore]` attributes
- [ ] Source generator emits extension methods on `ISheet`/`IStreamingSheet` (no `IRowMapper<T>` indirection; no `ISheet.AddRow<T>` method — types without `[Worksheet]` produce a compile-time error)
- [ ] `headerRow` is `int? = 1` (`null` = no header)

**Validation & errors**
- [ ] `WorkbookException` hierarchy
- [ ] Sheet name validation + `SanitizeSheetName`
- [ ] Cell address validation
- [ ] Bounded-resource enforcement on open

**Quality / infra**
- [ ] `<Nullable>enable</Nullable>`, zero warnings
- [ ] 100% XML doc coverage gate
- [ ] Public API snapshot tests
- [ ] Golden-file test corpus
- [ ] Round-trip tests
- [ ] Round-trip **preservation** test: open a workbook containing pivot caches, conditional formatting, custom XML, and threaded comments; modify one cell; save; assert unmodeled parts are bit-identical (§7.7 — v1.0 ship-blocker)
- [ ] Concurrent-mutation test: two threads mutating the same workbook produce `InvalidOperationException`, not corruption (decision #43)
- [ ] Use-after-dispose test: every public type throws `ObjectDisposedException` after `Dispose()` (decision #42)
- [ ] Benchmark suite vs NPOI / EPPlus / ClosedXML
- [ ] Source Link, deterministic builds, symbol packages, signed assemblies
- [ ] Cookbook samples project with the 13 recipes listed in design §8.1, each backed by a golden-file test
- [ ] Manual smoke-test checklist documented

### v1.1 — Common asks (target: TBD)

- [ ] Rich text in cells
- [ ] Excel Tables (`ListObject`)
- [ ] Image embedding (PNG / JPEG)
- [ ] Sheet protection
- [ ] Workbook protection
- [ ] Data validation (lists, ranges, custom)
- [ ] AutoFilter
- [ ] Named / reusable styles API
- [ ] Custom type converters for `Rows<T>`

### v2.0 — Advanced styling & charts (target: TBD)

- [ ] Conditional formatting
- [ ] Themes / theme colors
- [ ] Charts
- [ ] Modern threaded comments
- [ ] `.xlsm` passthrough
- [ ] Drawings / shapes
- [ ] Split panes
- [ ] Grouping / outlining
- [ ] Sorting helpers
- [ ] Streaming read
- [ ] Roslyn analyzers
- [ ] Breaking-change review and migration guide

### v3.0 — Power features (target: TBD)

- [ ] Pivot table writing
- [ ] Sparklines
- [ ] File-level encryption / password protection
- [ ] LINQ provider over sheets
- [ ] Breaking-change review and migration guide

## Explicit non-goals (forever)

- `.xls`, `.doc`, `.ppt`, `.xlsb` formats
- Formula evaluation (Excel does this)
- Pivot table *reading*
- Macros / VBA
- `netstandard2.0` / .NET Framework support (revisit only on explicit internal request)
