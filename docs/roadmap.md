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
| Split panes                             |      |      | [x]  |      |       |
| Hidden rows / columns / sheets          | Yes  |      |      |      |       |
| Auto-size columns                       | Yes  |      |      |      |       |
| Grouping / outlining                    |      |      | [x]  |      |       |
| Named ranges                            | Yes  |      |      |      |       |
| Tables (`ListObject`)                   |      | Yes  |      |      |       |
| **Data features**                       |      |      |      |      |       |
| Formulas (write only — Excel evaluates) | Yes  |      |      |      |       |
| Formula evaluation                      |      |      |      |      | No    |
| Data validation (drop-downs, ranges)    |      | Yes  |      |      |       |
| Conditional formatting                  |      |      | Yes  |      |       |
| AutoFilter                              |      | Yes  |      |      |       |
| Sorting                                 |      |      | [x]  |      |       |
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
| `.xlsm` (macro-enabled, passthrough)    |      |      | [x]  |      |       |
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
| Native AOT compatible                   |      |      |      |      | Deferred† |
| Trim compatible                         |      |      |      |      | Deferred† |

† AOT/Trim are **deferred, not refused.** Spike 4 (see
`spikes/results/spike-4-aot-trim.md`) measured both as runtime-incompatible
with NPOI 2.7.3 — `XSSFWorkbook` initialization throws `POIXMLException`
under both `PublishAot=true` and `PublishTrimmed=true`. The facade layer
itself is AOT-clean; the engine is not. The build-time MSBuild guards
`NXLS0100` / `NXLS0101` prevent silent runtime failures by failing the
build loudly.

Two paths to promote these matrix cells to `Yes`:
- **NPOI 3.x** removes its `System.Xml.Serialization` /
  `System.Reflection.Emit` dependencies (no announced date). Quarterly
  re-spike per `docs/scheduled-spikes.md` Spike 4-Q.
- **Native OOXML engine** (`docs/long-term.md`) implemented from
  scratch with AOT-clean design from day one. Long-term R&D track;
  not v1.0 work.

Either resolution lifts the deferral. The matrix is updated when one
of them lands, not when one of them is decided.

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

**Pre-implementation spike** (precondition to locking the design) — **complete 2026-05-15**:
- [x] Style-dedup feasibility benchmark — see `spikes/results/spike-1-style-dedup.md`. Design §4 perf row replaced (capacity + throughput, not overhead %).
- [x] Streaming-write back-pressure measurement on a 1M-row workload — see `spikes/results/spike-2-streaming-back-pressure.md`. In-memory target lowered to 30k rows; streaming on track for 1M.
- [x] Async wrapping cost — see `spikes/results/spike-3-async-wrapping-cost.md`. `Task.Run` wrapping is free or net-positive. Decision #5 confirmed.
- [x] **AOT / trim posture** — see `spikes/results/spike-4-aot-trim.md`. Both AOT and trim fail at runtime against NPOI 2.7.3. Matrix marked `No` for v1.0.

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
- [ ] `.github/workflows/` (S17): build, test, golden-file tests, public-API snapshot, benchmarks, pack, publish
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
- [ ] Headless-Linux test: `AutoSizeColumn` produces correct results when `libgdiplus` + a fallback font are installed; produces `MissingFontException` with installation guidance when they are not (decision I3)
- [ ] A1 parser test suite: every accepted form in design §6.10 canonicalizes correctly; every rejected form throws `InvalidCellAddressException` with a useful message
- [ ] `[Worksheet]` source-gen diagnostic catalog test: each of `NXLS0001`–`NXLS0006` is emitted on its trigger case (design §6.12)
- [ ] Benchmark suite vs NPOI / EPPlus / ClosedXML
- [ ] Source Link, deterministic builds, symbol packages, signed assemblies
- [ ] Cookbook samples project with the 13 recipes listed in design §8.1, each backed by a golden-file test
- [ ] Manual smoke-test checklist documented

### v1.1 — Common asks (target: tag pending release-PR + review pass)

Feature work landed 2026-05-22 across 10 slice commits (`e576b27` →
`a0c9acb`). Design decisions I-50…I-59 capture the surface shape +
rationale for each. The one remaining v1.1 item — the fuzz harness —
is in flight as part of the post-features push.

- [x] Rich text in cells — `e576b27` (I-50)
- [x] Excel Tables (`ListObject`) — `8c0aef5` (I-51)
- [x] Image embedding (PNG / JPEG) — `4b5ae48` (I-52)
- [x] Sheet protection — `49ed280` (I-53)
- [x] Workbook protection — `3861ac4` (I-54)
- [x] Data validation (lists, ranges, custom) — `a90e780` (I-55)
- [x] AutoFilter — `ae8a481` (I-56)
- [x] Named / reusable styles API — `f796aa9` (I-57)
- [x] Custom type converters for `Rows<T>` — `efe0092` (I-58)
- [x] `WorkbookOptions.StrictConcurrencyDetection` — `a0c9acb` (I-59).
      Opt-in real-lock mode replacing the opportunistic reentry counter
      (decision #43). Default false (single-threaded callers don't pay
      the lock cost); when true, takes a per-workbook Monitor lock on
      every mutating path.
- [x] **Fuzz harness for the open path** — landed post-v1.1-features as `tests/NetXlsx.Fuzz/` (xUnit-based corpus harness, not SharpFuzz — see `implementation-notes.md` for rationale). Found and fixed `IndexOutOfRangeException` leak from NPOI's parser on first run (decision I-60). Original entry:
      design's "bounded-resource parsing" claim (decision #20,
      `ReadMaxSheets` + `ReadMaxUncompressedBytes`) is enforced by a
      handful of explicit checks. A fuzz harness against
      `Workbook.Open` / `OpenAsync` with mutated zip + XML inputs
      would surface edge cases the explicit checks don't cover (XML
      expansion bombs, malformed relationships, OOXML schema
      violations that NPOI handles silently, etc.). High-leverage for
      the security posture; low LOC investment. Lives under
      `tests/NetXlsx.Fuzz/` as a separate project so it can be opt-in
      in CI (long-running on a nightly cadence) without blocking
      every PR.

### v1.2 — Close v1.1 deferments + post-tag polish (target: TBD)

v1.1 shipped 10 surface slices with five deliberate deferments — each
one was the right scope-control call at slice time, but the resulting
"reach through `.Underlying`" gaps are exactly what v1.2 closes. Plus
the one substantive open item from the v1.1 external review pass
(ISheet.cs SRP pressure at 888 LOC).

- [x] **`ISheet.RemoveTable(ITable)`** — landed (I-63). NPOI 2.7.3's `XSSFSheet` has
      no `RemoveTable` method; removal requires three coordinated
      mutations (drop the `<tablePart>` from `CT_Worksheet.tableParts`,
      remove the package relationship, update the sheet's part-loaded
      tables dictionary). Was deferred at v1.1 slice 2 — no observed
      user demand, and the three-step dance is fragile against NPOI
      internals. v1.2 either implements it carefully against 2.7.3
      or skips to NPOI 3.x (if the August re-checks unblock the bump).
- [x] **Per-column totals row on `ITable`** — landed (I-64). `ITable.HasTotalsRow` was
      read-only in v1.1 because adding a totals row requires per-column
      `SubTotalFunction` selection (Sum / Avg / Count / Min / Max /
      StdDev / Var / Custom). A reasonable v1.2 surface: extend the
      table-column model with a small enum + per-column setter.
- [x] **`IWorkbook.ProtectWithPassword(...)` — workbook-level password** — landed (I-65). Originally proposed as a `password:` argument on existing `Protect`; landed as a separate method to avoid call-site ambiguity + `RS0027`.
      NPOI 2.7.3 does not expose workbook-password APIs directly; v1.1
      shipped flag-only structure/windows/revision locks. v1.2
      manipulates `CT_WorkbookProtection` directly to attach the
      password hash (same UX-guard-not-security caveat as sheet
      protection in I-53).
- [x] **Per-column filter criteria for `ISheet.SetAutoFilter`** — landed (I-66) for the custom-filter variant. Explicit-list (`In(...)`) + Top-N variants deferred to v1.3 (NPOI 2.7.3 doesn't surface those properties on `CT_FilterColumn`; would need XML-node-level workarounds). v1.1
      ships range-only AutoFilter. Excel's filter-criteria model is
      rich (text equals / contains / not-equals, top-N, color, date
      range, custom expression). The v1.2 surface needs a builder or
      record-based criteria type — non-trivial design work.
- [ ] **OOXML named-style table integration** — deferred to v1.3
      after scope assessment in v1.2. v1.1 named styles (decision I-57)
      are an in-process convenience — the name → style map is not
      rehydrated by `Workbook.Open`. The full integration requires
      both a write-side (allocate `CT_Xf` in `cellStyleXfs`, register
      `CT_CellStyle` with name + xfId on each `RegisterStyle` call)
      AND a read-side (new CT_Xf → `CellStyle` parser distinct from
      the existing `CellStylePool.ReadFromNpoi(ICellStyle)`). The
      paired implementation is one focused slice in its own right;
      v1.2 ships the other five v1.1 deferments first, v1.3 leads
      with this one.
- [ ] **`ISheet.cs` partial-class split** (v1.1 review item 2). The
      file is at 888 LOC and approaching SRP pressure as the v1.1
      surfaces (Tables, Pictures, Protection, AutoFilter, Validation)
      all added partial files alongside the original. The shape is
      well-established: `XssfSheet.Tables.cs`, `XssfSheet.Pictures.cs`,
      `XssfSheet.Protection.cs`, `XssfSheet.AutoFilter.cs`,
      `XssfSheet.Validation.cs` already exist on the impl side. Mirror
      that on the public interface side — carve `ISheet` into
      `ISheet.cs` (core) + `ISheet.Tables.cs` (partial extension) +
      similar. Zero behavioral change; reduces future-contributor
      friction.
- [ ] **Reactive items from v1.1 usage feedback.** This bucket fills
      as the released v1.1.0 sees real-world use. Anything that's a
      bug fix lands in v1.1.x (patch); anything that's a missing
      capability lands here.

**Process note for v1.2.** Per the v1.1 working agreement, each
slice surfaces a design row (I-63, I-64, …) and a per-slice commit
with the standard checklist. The "go down the list" pattern from
v1.1 carries forward.

### v1.3 — OOXML named-style integration + AutoFilter follow-ups (target: TBD)

v1.2 deferred two items after in-flight scope assessment. v1.3 is
the natural home for both.

- [x] **OOXML named-style table integration** — landed (I-67).
      Both write side (CT_Xf in cellStyleXfs, CT_CellStyle in
      cellStyles) and lazy-rehydrate read side. Reaches across
      NPOI's internal `PutCellStyleXf` / `GetCellStyleXfAt` /
      `PutCellXf` via the centralized reflection in
      `Internal/NpoiInternals.cs`.
- [x] **`FilterCriteria.In(...)` explicit-value list filter** —
      partial landing (I-68). v1.3 ships 1–2-value support via the
      existing customFilters infrastructure (the same OR-joined-
      equality path Excel uses for short lists). 3+ values throw
      `NotSupportedException` until NPOI 3.x or a custom XML-
      emission workaround lifts the gap. API surface lifts
      cleanly when that happens — no caller-code change.
- [ ] **`FilterCriteria.Top(n)` / `BottomPercent(...)` Top-N filter** —
      deferred indefinitely; **blocked on NPOI 3.x**. Unlike `In(...)`
      (slice 2), Top-N has no customFilters fallback — the OOXML
      `<top10>` element is fundamentally a different filter type that
      can't be approximated with operator+value conditions. Adding the
      method surface today would be a footgun (always throws). Re-add
      when the upstream gap closes. Tracked here so it doesn't get
      lost; not a v1.3 deliverable.
- [ ] **Reactive items from v1.2 usage feedback.**

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

## Process rules

These govern *how* the roadmap is executed. They are referenced from `docs/design.md §3.1` decisions I19–I22.

### Tag cadence

- `v0.x.y` tags are cut throughout scaffold and v1.0 implementation. Internal consumers may take preview dependencies on any `v0.x.y` with the understanding that the API surface is still mutable.
- `v1.0.0` is cut only after every item in the v1.0 Definition of Done passes, including all four pre-impl spikes.
- After `v1.0.0`, the public API is locked per the SemVer policy in `design.md §3 #23`.

### Public-API snapshot transitions

- The PR that cuts a tagged release also moves the contents of `PublicAPI.Unshipped.txt` into `PublicAPI.Shipped.txt` for every project.
- Between releases, all additions go to `Unshipped.txt`. The CI PR check enforces this — no addition to `Shipped.txt` outside a release PR is permitted.

### Release-PR checklist (template, applies to every tagged release)

Every release-PR (v1.0, v1.1, v2.0, …) must include all of the
following as discrete commits or coherent stacked diffs — none
silently. Sections marked **(when applicable)** apply only to
releases that perform the action; skip cleanly when not.

- [ ] **PublicAPI flip.** Move every entry from `PublicAPI.Unshipped.txt`
      into `PublicAPI.Shipped.txt` for every library project. After
      the flip, `Unshipped.txt` contains only the `#nullable enable`
      header line.
- [ ] **`PublicApiSnapshotTests` baseline reconciliation.** Verify
      the expected baseline still matches by running the test.
- [ ] **TFM drops (when applicable).** When this release drops one
      or more TFMs per decision I24, remove them from
      `Directory.Build.props` `TargetFrameworks`, `global.json`
      rollForward target if relevant, every workflow's `setup-dotnet`
      block, and any spike/benchmark csproj that pinned the TFM.
- [ ] **CHANGELOG release entry** with explicit `### BREAKING CHANGES`
      section at the top when this release has any breaks. Include
      migration guidance and a reference to the relevant decision
      number(s) (e.g. I24 for a TFM drop). Rename the `[Unreleased]`
      header to `[X.Y.Z] — YYYY-MM-DD` and add a fresh empty
      `[Unreleased]` above it.
- [ ] **Doc-count sweep.** README, `~/dev/prompts/netxlsx-continuation.md`,
      and the upcoming CHANGELOG entry should all carry the
      post-release test totals and TFM counts. Historical CHANGELOG
      entries stay at the count they were written under.
- [ ] **Named ship-blockers landed.** Every checkbox in this
      release's per-version section above must be green. The gate is
      "all the boxes pass on the release-PR CI run" — no aspirational
      tags.
- [ ] **Version tag.** Tag `vX.Y.Z` from the merge commit (or the
      latest fix commit if release-PR validation surfaces a workflow
      bug — see the v1.0 retrospective entry in `implementation-notes.md`).
      MinVer picks up the tag for the NuGet package version. The
      release workflow fires on tag push and packs + publishes if
      `NUGET_API_KEY` is set.
- [ ] **`NUGET_API_KEY` secret** decision: either confirm it's set
      in repo Settings → Secrets and Actions before tag push (push
      lands synchronously with the tag), or accept that this release
      ships as a GitHub-only release with the .nupkg available from
      the GH Release page. Both are valid; pick deliberately.

#### v1.0 retrospective (lessons for future release PRs)

The v1.0.0 release-PR run surfaced two workflow bugs that weren't
visible until the release workflow ran end-to-end for the first time:

1. `release.yml`'s `Test (release smoke)` step ran ALL tests,
   including the `HeadlessNoFonts`-trait test that requires a
   font-less environment. The release runner installs `libgdiplus`
   + fonts (so AutoSize succeeds in the normal test paths), so the
   strict failure-path test fired and the workflow exited 1. Fix:
   `--filter "Category!=HeadlessNoFonts"` in the test invocation
   (matches `ci.yml`).
2. `NetXlsx.SourceGen` was being packed as a separate NuGet
   package, producing NU5128 ("declared netstandard2.0 deps but no
   `lib/netstandard2.0`") which became a build error under our
   `TreatWarningsAsErrors=true`. Fix: bundle the SourceGen dll into
   `NetXlsx.nupkg` under `analyzers/dotnet/cs/` via a target in
   `NetXlsx.csproj`; mark `NetXlsx.SourceGen.csproj` with
   `IsPackable=false`. Single package, no separate SourceGen
   .nupkg.

Both fixes landed before v1.0.0 was re-tagged. Future release PRs
should verify the release workflow on a dry-run (workflow_dispatch
or against a `v0.0.0-rc.N`-style throwaway tag) before tagging the
real version, to avoid the delete-and-re-tag dance.

### Spike-failure handling

- If a pre-impl spike misses its stated target, the project owner chooses among:
  1. Revise the target — record the new target as a design-doc revision **before** any code lands.
  2. Implement a workaround — record the workaround and any new caveats in `docs/design.md`.
  3. Descope the feature — move it out of v1.0 in this file.
- The outcome is recorded explicitly. Silent target erosion is not permitted.

### TFM additions

- New TFMs (`net10.0+`) are added in the next **minor** release after the TFM reaches GA. Never in a patch release. Never silently.
- The roadmap matrix's Platform rows are the source of truth — any addition is a roadmap PR first.

### Test fixture provenance

- Hand-crafted fixtures (built with Excel) are committed as binary `.xlsx` under `tests/NetXlsx.GoldenFiles/fixtures/`. Each has an entry in `tests/NetXlsx.GoldenFiles/README.md` documenting how it was produced.
- Generated fixtures live alongside a `.gen.cs` file that produces them on demand. CI verifies that the committed binary matches the generator output, modulo normalized non-deterministic OPC parts:
  - `docProps/core.xml` `created` and `modified` timestamps → normalized to a fixed sentinel before comparison
  - `docProps/app.xml` application name and version → normalized
  - ZIP central-directory entry ordering → sorted by part name before comparison
  - ZIP per-entry timestamps → zeroed
  
  Any other byte difference is a failure to investigate; it usually indicates an NPOI behavior change that warrants a workaround note in `docs/npoi-workarounds.md`.

## Explicit non-goals (forever)

- `.xls`, `.doc`, `.ppt`, `.xlsb` formats
- Formula evaluation (Excel does this)
- Pivot table *reading*
- Macros / VBA
- `netstandard2.0` / .NET Framework support (revisit only on explicit internal request)
