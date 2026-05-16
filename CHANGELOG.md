# Changelog

All notable changes to NetXlsx are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
per `docs/design.md §3 #23`. Pre-1.0 minor versions may include breaking
changes (decision I19).

## [Unreleased]

### Design-doc sync (2026-05-16)
Migrated three implementation-time additions from `implementation-notes.md`
up into `design.md` proper, per the methodology rule that load-bearing
API decisions land in the design, not in parallel notes. Audited via
comparison against the locked design surface; found broader drift than
the initial "one method family" estimate.

- §6.3 `ISheet` — added `IRow AppendRow()` to the regular sheet. (Was
  previously only on `IStreamingSheet` in the design; the v0.3.0 IRow
  slice added it to regular `ISheet` after the TabularExport recipe
  demanded a "find the next free row" idiom.)
- §6.5 `IRow` — added 10 fluent `Set(int col, T value)` overloads
  (string / bool / int / long / double / decimal / DateTime / DateOnly /
  TimeOnly / TimeSpan). Plus `ICell Cell(int col)` method form. These
  are the fluent setters that power `sheet.AppendRow().Set(1, x).Set(2, y)…`.
  The recipe-driven motivation is recorded inline with the design rows.
- §6.4 `ICell` — added `SetNumber(int)` and `SetNumber(long)` overloads
  resolving the literal-`42` ambiguity surfaced by the v0.2.0 cookbook.
- §9 substrate decisions S27 / S28 / S29:
  - S27: AOT/trim consumer build guard via `buildTransitive/NetXlsx.targets`
    emitting `NXLSAOT001/2` MSBuild errors.
  - S28: `.editorconfig` analyzer suppressions (`CA1716`, `CA1720`,
    `RS0026`) with rationale.
  - S29: lazy per-workbook date/time format-style cache as an interim
    until the full §4 style pool lands.
- §4 perf targets — note explaining the S29 interim cache and how it
  composes with the eventual full pool.

No code changes; this commit closes the design-vs-implementation drift
the v0.3.x slices accumulated. Implementation-notes retains the *story*
of how the decisions evolved (the recipe-driven motivation, the
literal-42 ambiguity, etc.); design.md now holds the *current state*.

### Added (since v0.3.0)
- **Date / time / duration on `ICell`** — `SetDate(DateTime)`,
  `SetDate(DateOnly)`, `SetTime(TimeOnly)`, `SetDuration(TimeSpan)`
  plus matching `GetDate()`, `GetDateOnly()`, `GetTime()`,
  `GetDuration()`. `IRow.Set` gains the four corresponding fluent
  overloads. Decisions honored:
  - I15: negative `TimeSpan` throws `ArgumentOutOfRangeException`.
  - I17: `DateTime.Kind` stored verbatim; reads always return
    `Kind = Unspecified`.
  - I-18 / I-19 / §7.9: default number formats applied lazily per
    workbook — `yyyy-mm-dd hh:mm:ss` for `DateTime`, `yyyy-mm-dd`
    for `DateOnly`, `h:mm:ss` for `TimeOnly`, `[h]:mm:ss` for
    `TimeSpan` (elapsed time). Explicit user-set styles are preserved
    (decision I-18).
  - §7.9: `GetTime()` returns `null` for fractional-day values outside
    `[0, 1)`; `GetDuration()` accepts any numeric cell.
- **Generator scope expanded.** `DateTime` / `DateOnly` / `TimeOnly` /
  `TimeSpan` properties now compile cleanly on `[Worksheet]` types
  (no more `NXLS0006`). `Guid` still trips `NXLS0006` — its
  setter overload is a separate future slice.
- **Cookbook recipe 12 — `TimeAndDuration`.** Demonstrates each
  date/time/duration kind with its default format, including the
  elapsed-time format that renders `26h` as `26:00:00` rather than
  modulo-24h `02:00:00`.

### Added (since v0.2.0)
- **`IRow` surface + `ISheet.AppendRow` / `Row(int)` / `[r,c]` indexer.**
  Real row API per design §6.4-§6.6. Fluent `IRow.Set(int col, T)` for
  every scalar kind (string, bool, int, long, double, decimal). The
  TabularExport recipe rewrite removes the v0.2.0 per-cell string
  arithmetic in favor of `sheet.AppendRow().Set(1, x).Set(2, y)...`.
- **`ICell.SetNumber(int)` and `SetNumber(long)` overloads.** Resolves
  the literal-`42` ambiguity surfaced by the v0.2.0 cookbook.
- **`[Worksheet]` generator emits real bodies for AddRow/AddRows.**
  `[Obsolete(error: true)]` removed from the write methods; bodies
  call `ISheet.AppendRow().Set(col, value)` per-property. Property
  types map to the right `IRow.Set` overload via the generator's
  `FormatSetCall` helper (handles narrowing/widening casts to
  disambiguate when the property type isn't a direct overload match).
  `ReadRows` still carries `[Obsolete(error: true)]` — the read-side
  typed-mapping slice is next.
- **Cookbook recipe 3 — `TypedExport`** (`SalesRecord` + source-gen).
  Demonstrates the `[Worksheet]`-driven write path end-to-end. The
  `[Worksheet(Visibility = WorksheetVisibility.Public)]` form is used
  so external consumers (golden-file tests) can call `sheet.AddRow(record)`.
- Generator `IsSupportedPropertyType` tightened for v0.3.x scope:
  `DateTime`, `DateOnly`, `TimeOnly`, `TimeSpan`, `Guid`, and any
  `Nullable<T>` value type now trip `NXLS0006` honestly. They'll
  pass when the corresponding `ICell.SetDate/SetTime/SetDuration`
  overloads land.
- `WorksheetProperty` model gains `UnderlyingSpecialType` (Roslyn
  enum, pipeline-cache-safe) so emit-side casts switch on a stable
  enum instead of a format-dependent type string.
- 11 new tests covering `IRow` / `Sheet[r,c]` / generator emit shape
  and the `TypedExport` recipe round-trip. Cookbook executable now
  dispatches `hello-workbook`, `tabular-export`, `typed-export`.

### Added (since v0.1.0)
- **v0.2.0 vertical slice — first working `.xlsx` round-trip.**
  - `Workbook.Create()` returns a real `IWorkbook` over `XSSFWorkbook`.
  - `Workbook.Open` / `OpenAsync` (path and stream forms) with stream-
    position-zero + seekable validation (decisions #50, I14) and
    `MalformedFileException` for non-.xlsx content (#51).
  - `IWorkbook.AddSheet` (case-insensitive uniqueness — decision #41,
    name validation per Excel rules), `Sheets` indexer (string and int),
    `TryGetSheet`, `SaveAsync` (Task.Run-wrapped per #30 / spike 3).
  - `IWorkbook : IDisposable` with `ObjectDisposedException` on
    use-after-dispose, safe double-dispose (decision #42).
  - `ISheet["A1"]` returns a materialized `ICell` even for never-written
    addresses (decision #40 — auto-blank).
  - `ICell.SetString/SetNumber(double|decimal)/SetBool/Clear` + the
    typed `GetString/GetNumber/GetBool` readers and `Kind` classifier.
    `decimal` writes documented as IEEE-754 lossy per #36 / §7.4.
    `GetString` follows §7.10's per-kind formatting rules including
    Excel error-code text and `"TRUE"`/`"FALSE"` invariant for bool.
  - `CellAddress.Parse` / `TryParse` / `Format` — A1 grammar per §6.10
    (single-cell form only; range parsing lands with `IRange` later).
    Accepts `A1`, case-insensitive variants, and `$A$1`/`$A1`/`A$1`
    (`$` stripped). Rejects `Sheet1!A1`, ranges (`A1:C10`, `A:A`,
    `1:1`), and overflow past `XFD` / row 1,048,576.
  - Exception hierarchy: `WorkbookException` + `InvalidCellAddressException`,
    `SheetNameException`, `MalformedFileException`.
  - `CellKind` enum: `Empty / String / Number / Date / Bool / Formula / Error`.
- Deleted `Placeholder.cs` — first real types replaced it as planned.

### Cookbook recipes 1-2 (executable + golden-file tested)
- `HelloWorkbook` recipe — string / number / decimal / bool round-trip
  through the v0.2.0 cell API. Tested via both NetXlsx reopen *and*
  direct NPOI read-back (catches writer-only-bug class of issues).
- `TabularExport` recipe — writes N records as rows of cells using the
  current `sheet["A{r}"]` indexer. Deliberately clunky at scale; the
  awkwardness is documented in the recipe itself as the load-bearing
  motivation for the next slice's `IRow` API.
- Cookbook executable dispatches recipes by name:
  `dotnet run --project samples/NetXlsx.Cookbook -- hello-workbook /tmp/out.xlsx`
- `tests/NetXlsx.GoldenFiles/Recipes/` invokes the same recipe
  classes the executable does (project reference, not code duplication).
- Recipes added 5 new tests (2 HelloWorkbook + 3 TabularExport).
  Total: 70 tests on each TFM.

### Ergonomics gap surfaced (logged for next slice)
- `cell.SetNumber(42)` is ambiguous between the `double` and `decimal`
  overloads — a real call-site footgun for integer-literal cases. Likely
  resolution: add `SetNumber(int)` / `SetNumber(long)` overloads on
  `ICell`. Logged in implementation-notes for the IRow slice.

### Known limitations of v0.2.0 vertical slice
- No `[r,c]` cell indexer yet (only `["A1"]`); arrives with the row /
  column / range API.
- No `IRow` interface yet; the typed-mapping source generator's emitted
  `AddRow` / `AddRows` / `ReadRows` extensions remain `[Obsolete(error)]`
  pending a follow-up commit that wires their bodies to the new `ISheet`.
- No concurrent-mutation detection yet (decision #43); NPOI is not
  thread-safe and we don't lock — documented now, enforced later.

### Added
- Initial project scaffold (Directory.Build.props, Directory.Packages.props,
  nuget.config, .editorconfig, LICENSE, CODEOWNERS, build scripts,
  TeamCity DSL placeholder, source project skeletons, test/benchmark/
  sample/golden-file project skeletons, public-API snapshot files).
- Strong-name key (`netxlsx.snk`) generated and committed.
- Public marker attributes: `[Worksheet]`, `[Column]`, `[Ignore]`, plus
  `WorksheetVisibility` enum. Empty `ISheet` / `IWorkbook` marker
  interfaces (members per design §6.4 land in milestone 2).
- `WorksheetGenerator` (`IIncrementalGenerator`): scans the current
  compilation for `[Worksheet]` types, emits `{Type}_SheetExtensions`
  with `AddRow` / `AddRows` / `ReadRows` extension methods on `ISheet`.
  Bodies throw `NotImplementedException` pending milestone-2 `ISheet`
  implementation.
- Full diagnostic catalog `NXLS0001`–`NXLS0006` per design §6.12,
  with `AnalyzerReleases.{Shipped,Unshipped}.md` release tracking.
- 12 source-generator tests (one per diagnostic ID including the
  record-primary-constructor satisfaction case, plus 4 emission tests
  covering valid output, visibility opt-in, fatal-diagnostic short
  circuit, and the cross-assembly-ignored contract from I5).
- All four pre-implementation spikes complete; results captured under
  `spikes/results/`. Three triggered design revisions:
  - Spike 4 (AOT/trim): both `PublishAot` and `PublishTrimmed` fail at
    runtime against NPOI 2.7.3. Roadmap matrix marked `No`; decision I2
    updated from "TBD pending" to measured outcome.
  - Spike 2 (streaming back-pressure): in-memory 100k-row target missed
    by ~2×. Target lowered to 30k rows; streaming recommended above that.
    Streaming sustained flat ~70 MB ΔGC at 500k rows.
  - Spike 1 (style dedup): "10%/30% overhead vs raw NPOI" framing was
    measuring a phantom — raw NPOI hits its style cap before completing
    any styled workload of meaningful size. Replaced with absolute
    capacity + throughput targets.
  - Spike 3 (async wrapping): `Task.Run` wrapping is free or net-positive
    at every size. Decision #5 stands without revision.

### Known scaffold placeholders (TODOs)
- `nuget.config` — internal feed URL pending.
- `Directory.Build.props` `RepositoryUrl` — github.com/jkindrix pending.
- `CODEOWNERS` — owning team identifiers pending.
- Source Link package is not wired; depends on github.com/jkindrix choice.

## [0.1.0] — TBD

The first tagged scaffold release. Cut once the placeholders above are
filled, the first green CI run completes, and the four pre-implementation
spikes (style-dedup, streaming back-pressure, async wrapping cost, AOT/trim
posture — see `docs/roadmap.md`) are scheduled or running.
