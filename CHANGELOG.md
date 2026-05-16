# Changelog

All notable changes to NetXlsx are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
per `docs/design.md §3 #23`. Pre-1.0 minor versions may include breaking
changes (decision I19).

## [Unreleased]

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
