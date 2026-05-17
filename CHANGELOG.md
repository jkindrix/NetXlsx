# Changelog

All notable changes to NetXlsx are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
per `docs/design.md §3 #23`. Pre-1.0 minor versions may include breaking
changes (decision I19).

## [Unreleased]

### v0.9 — Streaming write (`IStreamingWorkbook`, `Workbook.CreateStreaming`)
Lands the biggest remaining v1.0 ship-blocker — write-side streaming via
NPOI's SXSSF. Random-access write/read stays on `IWorkbook` /
`Workbook.Create`; bulk writes past ~30k rows now have a first-class
entry point that holds memory flat per spike 2.

Public surface (PublicAPI.Unshipped.txt: +63 entries):
- `Workbook.CreateStreaming(StreamingOptions? options = null) ->
  IStreamingWorkbook` — entry point per design §6.1.
- `IStreamingWorkbook : IDisposable, IAsyncDisposable` with
  `AddSheet`, `Save` (sync + async, stream + path), and
  `Underlying` returning `SXSSFWorkbook`.
- `IStreamingSheet` with `AppendRow()` / `AppendRow(int index)` —
  the latter enforces the append-only contract by throwing if
  `index <= last written`. `Underlying` returns `SXSSFSheet`.
- `IStreamingRow` with `Index`, indexers `[int]` / `[string]`,
  `Cell(int)`, seven `Set(int, T)` fluent overloads (string,
  double, decimal, int, long, bool, DateTime), and an explicit
  `Flush()` that delegates to `SXSSFSheet.FlushRows()`.
- `IStreamingCell` — **new, sibling to ICell** (design decision
  **I-49**, see implementation-notes). Has the value setters
  (`SetString`/`SetNumber`/`SetBool`/`SetDate`/`SetFormula`),
  `Style(CellStyle)`, `NumberFormat(string)`, address + `Kind`.
  No `Underlying` — NPOI's `SXSSFCell` doesn't inherit
  `XSSFCell`, so the `ICell.Underlying : XSSFCell` contract
  cannot be honored; consumers reach the raw cell through
  `IStreamingSheet.Underlying`.
- `WorkbookOptions` (property bag): `DisplayCulture`,
  `DateSystem`, `ReadMaxUncompressedBytes`, `ReadMaxSheets`,
  `MaxRowsPerSheet`, `MaxColsPerSheet`, `MaxCellTextLength`,
  `DefaultFontName`, `DefaultFontSize`. Defaults match Excel.
  v0.9 wires `StreamingOptions` properties only; the random-access
  side will pick up `WorkbookOptions` overloads in a follow-up.
- `StreamingOptions : WorkbookOptions` adds `RowAccessWindowSize`
  (default 100, NPOI default) and `CompressTempFiles` (default
  false).
- `DateSystem { Excel1900, Excel1904 }` per design §6.1.

Cookbook (v1.0 set is now **13 of 13**):
- **StreamingMillionRows** (recipe 9). Defaults to 250k rows × 20
  columns (CI-friendly); `Run(path, rowCount)` overload lets ops
  bump it up for a true perf check. Mixes int/double/string cells
  so it isn't a numeric-fast-path-only demo. Sized at 5,000 rows
  in the golden-file test for fast CI feedback.

Internal:
- New `Internal/Sxssf{Workbook,Sheet,Row,Cell}.cs` wrappers.
  `SxssfWorkbook` owns both the SXSSF wrapper *and* the underlying
  XSSF base so the style pool (which needs an `XSSFWorkbook`) can
  be reused unchanged.
- `SxssfCell.Style` merges overlay-non-null over the cell's
  current style and routes through the same `CellStylePool` as
  random-access. Reverse-lookup from NPOI's `ICellStyle` to a
  `CellStyle` record is only fully reliable for `NumberFormat`
  (the streaming code doesn't index font/fill/border tables);
  documented as a known weaker-merge corner in implementation-notes.

Design + notes sync:
- `docs/design.md §6.3` rewritten with the `IStreamingCell` split
  and decision I-49 reference.
- `docs/implementation-notes.md` carries the full explanation:
  why ICell couldn't be reused, what `IStreamingSheet.Underlying`
  buys callers, and the merge-semantics caveat.

Tests (+20 unit + 1 golden = +21):
- `StreamingWorkbookTests` covers entry point + lifecycle (type
  separation from `IWorkbook`, default and explicit window size,
  double-dispose safe), `AddSheet` (name validation, dup
  rejection case-insensitive), append-only contract (start at 1,
  monotonic increment, explicit-index skip-forward, cannot
  revisit, grid-bound validation), cell-level write (all seven
  scalar `Set` overloads round-trip through Save→Open via the
  random-access reader; letter indexer; column-bound validation),
  `SaveAsync` round-trip, dispose-throws matrix, formula
  (round-trip + empty/null rejection), and `NumberFormat`
  surviving save-open.
- Cookbook golden-file test runs 5,000-row streaming write,
  checks file size + spot-checks header / first / middle / last
  / mixed-type cells.
- `PublicApiSnapshotTests` baseline extended with the six new
  public types (`DateSystem`, `IStreamingCell`, `IStreamingRow`,
  `IStreamingSheet`, `IStreamingWorkbook`, `StreamingOptions`,
  `WorkbookOptions`).

Cookbook is now **complete at 13 of 13 recipes** for the v1.0 set.

### v0.8.1 — Cookbook recipes 10, 11 (NPOIEscapeHatch, OpenEditSave)
Two more cookbook recipes — cookbook is now **12 of 13**, with only
`StreamingMillionRows` (recipe 9) remaining, gated on the streaming
write slice.

Recipes:
- **NPOIEscapeHatch** (recipe 10). Demonstrates the design's
  first-class-escape-hatch promise (decisions #1, #32). Builds a
  small data sheet through the normal facade, then reaches through
  `ISheet.Underlying` to set a print area, configure landscape +
  fit-to-1-page-wide page setup, write header/footer text, and
  repeat the header row on every printed page — all operations
  v1 deliberately doesn't model. The wrapper still owns the
  workbook lifecycle; the escape hatch is for incremental
  capability, not workaround.
- **OpenEditSave** (recipe 11). Builds an input file via raw NPOI
  carrying a custom OPC part (`/customXml/itemRecipe.xml`), then
  opens it through NetXlsx, mutates two cells (one append, one
  overwrite), and applies an identical style to two cells. The
  golden test asserts both preservation promises in a single run:
  §7.5 (style pool dedup — A1 and B1 share one NPOI
  `ICellStyle.Index`) and §7.7 (the custom OPC part round-trips
  byte-identical). Single self-contained recipe; no committed
  fixture required.

Tests (+2 golden-file): one per recipe; golden suite now 26 per TFM
(up from 24).

### v0.8 — Cookbook recipes 6, 7, 8 (Formulas, MultiSheet, HyperlinksAndComments)
Adds the three cookbook recipes that v0.7 unblocked. Cookbook is now
**10 of 13** recipes. Each recipe has a paired golden-file test in
`tests/NetXlsx.GoldenFiles/Recipes/` per the established pattern.

Recipes:
- **Formulas** (recipe 6 from design §8.1). A "quarterly sales" sheet
  with per-row `=B*C` subtotals plus a `=SUM`, `=AVERAGE`, and
  `=Total*0.07` tax line. Demonstrates both leading-`=` and bare-body
  forms; asserts no cached values are pre-computed (NPOI's
  `NumericCellValue == 0.0` on every formula cell).
- **MultiSheet** (recipe 7). Three sheets — `Data` (12 months of
  sales + region), `Lookup` (region code → name), `Summary` —
  with two workbook-scoped named ranges (`MonthlySales`,
  `RegionLookup`) wired into `=SUM(MonthlySales)`,
  `=AVERAGE(MonthlySales)`, `=MAX(MonthlySales)`, and a
  `=VLOOKUP(..., RegionLookup, 2, FALSE)`. Demonstrates the
  documentation value of named ranges — formulas read as
  `=SUM(MonthlySales)` rather than `=SUM(Data!B2:B13)`.
- **HyperlinksAndComments** (recipe 8). Four hyperlinks exercising
  every supported scheme (decision I13: `https://`, `mailto:`,
  `file://`, internal `#Sheet!Range`) plus three comments —
  two with the default `"NetXlsx"` author (decision I11) and
  one with an explicit `release-bot` override.

Cookbook program: recipes registered in `Program.cs` so they're
runnable via `cookbook formulas`, `cookbook multi-sheet`,
`cookbook hyperlinks-and-comments`. Help text's "Recipes (v0.2.0)"
header dropped — it was stale.

Tests (+3 golden-file): one round-trip golden test per recipe;
total golden-file suite is now 24 per TFM (up from 21).

### v0.7 sub-slice C — Cell annotations (`ICell.Comment` / `Hyperlink` + read-side accessors)
Final third of the v0.7 bundle. Closes the v1.0 ship-blocker rows for
cell-level comments and hyperlinks per design §3 #368–369. Realizes
decisions I11 (default comment author) and I13 (hyperlink
scheme-sniffing).

Public surface (PublicAPI.Unshipped.txt: +5 entries):
- `ICell.Comment(string text, string? author = null) -> ICell` —
  attaches a comment. Default author is `"NetXlsx"` per I11
  (avoids leaking `Environment.UserName`). Replacing a comment
  mutates the existing one in place (NPOI rejects creating a
  second comment on the same cell).
- `ICell.GetComment() -> string?` — comment body, or null.
- `ICell.GetCommentAuthor() -> string?` — author, or null.
- `ICell.Hyperlink(string target, string? display = null) -> ICell`
  — attaches a hyperlink. `target` is scheme-sniffed per I13:
  `http(s)://`, `mailto:`, `file://`, internal `#Sheet!Range`.
  Anything else (`ftp://`, `javascript:`, bare paths) throws
  `ArgumentException`. If `display` is supplied, the cell's
  displayed string is set to it; if not and the cell is empty,
  it falls back to the raw target; if the cell already has text,
  the text is preserved.
- `ICell.GetHyperlink() -> string?` — hyperlink address, or null.

Internal:
- `XssfCell.Comment` lazily creates the sheet's drawing patriarch
  + a small (2x2) client anchor on first use; subsequent calls
  mutate the existing `IComment` in place. `IComment.String`
  goes through the creation helper's `CreateRichTextString`.
- `XssfCell.Hyperlink` constructs an `XSSFHyperlink` with the
  scheme-sniffed `HyperlinkType` and assigns it via
  `XSSFCell.Hyperlink = ...` (NPOI handles the
  `SetCellReference` / `AddHyperlink` wire-up).
- Internal `#Sheet!Range` form strips the leading `#` for
  consistency with NPOI's `Document`-type storage; `GetHyperlink`
  returns the body verbatim.

Tests (+22):
- `CommentAndHyperlinkTests` covers: default author = "NetXlsx",
  explicit author, in-place replace, fluent chaining,
  null-text rejection, null-on-no-comment getters, Save→Open
  round-trip, supported-scheme acceptance theory (https, http,
  mixed-case, mailto, file), internal `#Sheet!Range` form,
  unsupported-scheme rejection theory (ftp, javascript, bare path,
  absolute path), null/empty target rejection, display-replaces-text,
  display-null-on-empty-cell falls back to target,
  display-null-on-populated-cell preserves text, and full
  Save→Open round-trip with both URL and mailto schemes.
- `DisposedWorkbookMatrixTests` (+5): adds Comment, GetComment,
  GetCommentAuthor, Hyperlink, GetHyperlink to the
  `CellOperations` matrix.

### v0.7 sub-slice B — Named ranges (`IWorkbook.AddNamedRange`, `NamedRanges`, `INamedRange`)
Second third of the v0.7 bundle. Lands the workbook-level named-range
contract from design §3 #212–213 and §6.2.

Public surface (PublicAPI.Unshipped.txt: +5 entries):
- New interface `INamedRange` with `Name`, `Formula`,
  `SheetScope` (string? — `null` == workbook-scoped per decision I9).
- `IWorkbook.AddNamedRange(string name, string formula, string?
  sheetScope = null) -> INamedRange` — single overload (I9).
  Leading `=` on the formula is stripped for consistency with
  `SetFormula`. Returns the created range so callers can chain
  property inspection.
- `IWorkbook.NamedRanges` — workbook-wide enumeration (scope-agnostic).

Validation:
- Null / empty `name` or `formula` rejected with
  `ArgumentNullException` / `ArgumentException`.
- `sheetScope` referencing an unknown sheet throws
  `SheetNameException`.
- Duplicate names rejected with `ArgumentException` (case-insensitive)
  per the NPOI 2.7.x constraint documented below.
- NPOI parse failures (invalid name text, cell-reference-style names
  like `R1`) wrapped in `ArgumentException` with the original
  preserved as `InnerException`.

NPOI quirk handled (see implementation-notes for full discussion):
NPOI 2.7.x rejects coexistence of a workbook-scope name and a
sheet-scope name with the same text, even though Excel itself
permits it. v1 enforces workbook-wide uniqueness regardless of
scope. Revisit if/when NPOI relaxes this.

Internal:
- New `Internal/XssfNamedRange.cs` — wraps NPOI's `IName`.
  `SheetScope` resolves the workbook's sheet index back to a name
  (returns `null` for index `< 0` == workbook scope).
- `XssfWorkbook` uses NPOI's `GetAllNames()` (the post-3.16
  replacement for the deprecated `GetNameAt(int)`).

Tests (+13):
- `NamedRangeApiTests` covers: workbook-scope and sheet-scope
  round-trip, leading-`=` strip, empty `NamedRanges` on fresh
  workbook, multi-range enumeration, null/empty validation,
  unknown-sheet-scope rejection, case-insensitive duplicate
  rejection at workbook scope, same-name-different-scope rejection
  (NPOI constraint), and Save→Open round-trip of a named range
  used in a cross-sheet `SUM(Sales)` formula.
- `DisposedWorkbookMatrixTests` (+2): `AddNamedRange` and
  `NamedRanges` added to the `WorkbookOperations` matrix.
- `PublicApiSnapshotTests` baseline now includes `INamedRange`.

### v0.7 sub-slice A — Formula API (`ICell.SetFormula` / `GetFormula`)
First third of the v0.7 "formula + named-range + annotation" bundle.
Closes the v1.0 ship-blocker row for write-side formula support;
reads were already covered (formula cells classified as
`CellKind.Formula`, with `GetString`/`GetNumber`/`GetBool`/`GetError`
all routing through `CachedFormulaResultType`).

Public surface (PublicAPI.Unshipped.txt: +5 entries):
- `ICell.SetFormula(string)` — stores a formula. Leading `=` is
  optional (`"=A1+B1"` and `"A1+B1"` both accepted). Empty body
  rejected with `FormulaException`. NPOI parse failures translated
  to `FormulaException` with the original exception preserved as
  `InnerException` so callers don't see the NPOI type leak through.
  Per decisions #46 / §7.8 the cached value is **not** pre-computed —
  Excel and other competent consumers recalculate on open.
- `ICell.GetFormula() -> string?` — returns the formula body with
  a re-attached leading `=`, or `null` for non-formula cells.
- New exception `FormulaException : WorkbookException` with the
  two-arg constructor pair (message-only / message-and-inner). Per
  design §6 the design.md already enumerated this exception; the
  implementation realizes it now.

Internal:
- `XssfCell.SetFormula` strips an optional leading `=` so callers
  can write either form. NPOI's `CellFormula` property expects the
  body without the `=`; we wrap the parse path in a try/catch that
  translates to `FormulaException`. No pre-computation hook is
  invoked — `XSSFFormulaEvaluator` is deliberately not touched.
- `XssfCell.GetFormula` reads `CellType.Formula` cells only; all
  others return `null`. Body is prefixed with `=` for round-trip
  symmetry with `SetFormula`.

Tests (+10):
- `FormulaApiTests` covers: leading-`=` round-trip, no-`=` form
  acceptance, null rejection, empty-body rejection (`""` and `"="`),
  garbage-body translation to `FormulaException`, GetFormula null
  on non-formula cells, sheet-qualified-reference round-trip,
  no-pre-computation assertion (verifies
  `CachedFormulaResultType == Numeric` with default `0.0`),
  SetFormula-replaces-prior-value semantics, Clear-after-SetFormula
  resets to Empty, and full Save→Open round-trip.
- `DisposedWorkbookMatrixTests` (+2): adds `SetFormula` and
  `GetFormula` to the `CellOperations` matrix.
- `PublicApiSnapshotTests` baseline now includes `FormulaException`.

Known flake (pre-existing, unrelated to this slice):
`WorkbookRoundTripTests.Concurrent_AddSheet_Throws_InvalidOperationException`
is racy — it asserts the reentry-counter (decision #43) fires under
contention but depends on actually observing a collision. The
detection is "best-effort by design (it's not a lock)" per the test's
own comment. Will tighten the race window in a follow-up (use a
`Barrier` to synchronize start of both mutator threads).

### v0.6 sub-slice C — Column API (`IColumn`, `ISheet.Column(...)`, AutoSize / Hidden / Width / SetDefaultStyle)
Final third of the v0.6 bundle. Closes the v1.0 ship-blocker rows for
column-level width control, hidden columns, default-style fan-out, and
`AutoSize` with explicit headless-Linux behavior.

Public surface (PublicAPI.Unshipped.txt: +21 entries):
- New interface `IColumn` with:
  - `Index` (1-based) and `Letter` (canonical, `1 → "A"`).
  - `Sheet` — owning sheet.
  - `Hidden` (bool, read/write) — maps to NPOI's
    `IsColumnHidden` / `SetColumnHidden`.
  - `WidthUnits` (double, read/write) and fluent `Width(double)`.
    Width is in Excel "character" units; NPOI's 256ths-of-a-character
    integer representation is hidden inside the wrapper. Setter
    rejects negative and NaN.
  - `AutoSize()` — sizes to fit populated contents. On headless
    environments without a usable font stack, throws the new
    `MissingFontException` with installation guidance for
    Debian/Ubuntu and Alpine (design decision I3). Translation
    covers `SixLabors.Fonts.*`, `System.Drawing.SystemFontsException`,
    `TypeInitializationException`, and font-related IO failures.
  - `ForEachPopulated(Action<ICell>)` — sparse top-to-bottom
    iteration over populated cells in this column (empties skipped).
  - `SetDefaultStyle(CellStyle)` — applies via `CellStylePool`
    so identical column-default styles share one NPOI
    `ICellStyle` index; delegates to NPOI's
    `SetDefaultColumnStyle` so new cells in the column inherit.
- `ISheet.Column(int)` and `ISheet.Column(string letter)` factories.
- New exception `MissingFontException : WorkbookException` with the
  standard four-constructor pattern (parameterless / message-only /
  inner-only / message-and-inner). Default message includes
  install commands and points callers at `IColumn.Width(double)`
  as the deterministic alternative.
- `CellAddress.ParseColumn(string)` / `TryParseColumn(string, out int)`
  / `FormatColumn(int)` — public letter ↔ index helpers, reusing
  the same parser as `ParseRange`'s whole-column shorthand path.
  `FormatColumn` is the single-letter form of `Format(row, column)`.

Internal:
- New `Internal/XssfColumn.cs` — `IColumn` implementation. AutoSize
  failure translation lives here, as a dedicated
  `IsFontFailure(Exception)` helper that walks the inner-exception
  chain so wrapped failures (`TypeInitializationException` →
  `FileNotFoundException(libgdiplus)`) still get classified.
- `XssfSheet` gains the two `Column(...)` overloads.

Tests (+30):
- `ColumnApiTests` (+16) covers: construction by index and letter
  (including `aa`, `$AB`, `XFD` and the `FormatColumn` round-trip),
  letter-form rejection of garbage (empty, digits, A1, > XFD),
  index bounds validation, width round-trip through NPOI's
  quantization, fluent return identity, negative/NaN rejection,
  Hidden round-trip, `SetDefaultStyle` pool-routed application,
  `ForEachPopulated` sparse ordering, no-op on empty columns,
  null-action rejection, `AutoSize` succeed-or-throw-MissingFont
  (both outcomes acceptable; silent failure is not), and full
  Save→Open round-trip of width / hidden / default-style.
- `DisposedWorkbookMatrixTests` (+13): adds two `ISheet.Column(...)`
  entries plus an entire `ColumnOperations` matrix asserting every
  `IColumn` member throws `ObjectDisposedException` after
  `Workbook.Dispose()`.
- `PublicApiSnapshotTests`: baseline now includes `IColumn`
  and `MissingFontException`.

### v0.6 sub-slice B — Range API (`IRange`, `ISheet.Range(...)`)
Second third of the v0.6 bundle. Introduces a first-class rectangular
range abstraction so callers can fill, style, merge, or clear an entire
block without iterating cell-by-cell.

Public surface (PublicAPI.Unshipped.txt: +13 entries):
- New interface `IRange : IEnumerable<ICell>` with:
  - `Address` — canonical A1 range string (e.g. `A1:C3`).
  - `FirstRow` / `LastRow` / `FirstCol` / `LastCol` — inclusive
    1-based bounds.
  - `Count` — dense coordinate count (`rows * cols`), not
    populated-cell count.
  - `Sheet` — owning sheet.
  - `EnumerateAll()` — dense iteration that materializes every
    cell in the rectangle (including empties). Default
    `GetEnumerator()` is sparse — only populated cells are yielded.
  - `Value(object?)` — bulk fill. Runtime-type-dispatched
    (`string`, `int`/`long`/`double`/`decimal`, `bool`, `DateTime`).
    `null` clears every cell. Unsupported types throw
    `ArgumentException`. Returns `this` for chaining.
  - `Apply(CellStyle)` — bulk style. Goes through `CellStylePool`
    so every cell in the rectangle shares a single `ICellStyle`
    by index. Returns `this`.
  - `Merge()` — convenience that delegates to
    `ISheet.MergeCells(Address)`. Returns `this`.
  - `ClearContents()` — clears each cell's value but preserves
    its style (mirrors Excel's "Clear → Contents" command).
    Returns `this`.
- `ISheet.Range(string a1Range)` — A1-form factory. Now accepts
  whole-row (`3:3`) and whole-column (`A:A`) shorthand;
  `CellAddress.ParseRange` expands these to the full sheet
  bounds (`CellAddress.MaxRow` / `MaxColumn`). Sub-slice A
  shipped explicitly *without* this expansion; sub-slice B
  enables it because `IRange` is the consumer that needs it.
- `ISheet.Range(int row1, int col1, int row2, int col2)` —
  coordinate-form factory. Bounds-checked; corner order is
  normalized so callers can pass corners in any order.

Internal:
- `CellAddress.ParseRange` extended with `TryParseColumnOnly`
  and `TryParseRowOnly` helpers; the whole-row/column branches
  expand directly to `(1, col, MaxRow, col)` /
  `(row, 1, row, MaxColumn)`.
- New `Internal/XssfRange.cs` — full `IRange` implementation.
  Sparse `GetEnumerator()` walks NPOI's physical rows/cells;
  dense `EnumerateAll()` materializes coordinates through the
  same `XssfSheet[row, col]` indexer that materializes on access
  (decision #40).
- `XssfSheet` gains a private `ValidateGridCoordinate` helper
  reused by both `Range(int,int,int,int)` and the indexer-side
  validation path.

Tests (+15):
- `RangeApiTests` covers: A1 and coordinate-form construction,
  inverted-corner normalization, bounds validation, single-cell
  ranges, whole-row/column expansion, runtime-type dispatch
  (`int`/`long`/`double`/`decimal`/`bool`/`DateTime`/`null` and
  unsupported-type rejection), `Apply` style-pool dedup,
  `Merge` delegation, sparse vs dense enumeration,
  `ClearContents` preserves style index, and full
  Save→Open round-trip.
- `CellAddressTests` — the previous "rejects A:A / 1:1" theory
  is replaced by a positive expansion theory; an
  invalid-shape rejection theory still guards malformed forms.
- `DisposedWorkbookMatrixTests` — adds `ISheet.Range(string)`
  and `Range(int,int,int,int)` to the sheet-level matrix, plus
  a new `IRange` operation matrix that asserts every `IRange`
  member throws `ObjectDisposedException` after `Workbook.Dispose()`.

Compatibility:
- No breaking change. Whole-row/whole-column A1 forms that
  previously threw `InvalidCellAddressException` now succeed
  in sub-slice B; this is a behavior expansion, not a contract
  break (the diagnostic was a v0.6-sub-slice-A placeholder
  documented as such in `CHANGELOG.md` sub-slice A entry).

### v0.6 sub-slice A — freeze panes + merge cells + hidden rows/sheets + gridlines
The first third of the v1.0 "range / freeze / merge / hidden / autosize"
bundle. Each member here is on `ISheet` or `IRow` directly — no new
interface required.

Public surface (PublicAPI.Unshipped.txt: +14 entries):
- `ISheet.FreezeRows(int)`, `FreezeColumns(int)`, `FreezePane(int, int)`.
  Negative arguments throw `ArgumentOutOfRangeException`. Internally
  these delegate through `FreezePane`; NPOI's argument order
  (`colSplit, rowSplit`) is reversed inside the wrapper so the public
  API reads `(rows, cols)`.
- `ISheet.MergeCells(string)`, `UnmergeCells(string)`, `MergedRanges`.
  - `MergeCells` parses the A1 range, checks for overlap with existing
    merges (throws `InvalidOperationException` per design §6.4), and
    falls through to NPOI's `AddMergedRegion`. 1×1 ranges are no-ops
    per decision I-38.
  - `UnmergeCells` removes the exact-matching merged region, or
    silently no-ops if no exact match exists (design §6.4).
  - `MergedRanges` returns canonical `A1:C3` strings.
- `ISheet.Hidden` (bool) — workbook-level sheet visibility. Maps to
  NPOI's `SheetVisibility.Hidden` ↔ `Visible`. `VeryHidden` (hidden
  from VBA) intentionally not modeled in v1; reach through `Underlying`.
- `ISheet.ShowGridlines` (bool).
- `IRow.Hidden` (bool) — maps to NPOI's `ZeroHeight`.
- `CellAddress.ParseRange(string)` returning a 4-tuple `(Row1, Col1,
  Row2, Col2)`. Accepts `A1:C3` and single-cell forms; normalizes
  inverted corners. Whole-row (`1:1`) and whole-column (`A:A`)
  expansion is explicitly *not* supported here — those forms ship
  with the `IRange` API in sub-slice B.
- `CellAddress.FormatRange(int, int, int, int)` returning canonical
  `A1:C3` (1×1 collapses to single-cell form `A1`).

Cookbook recipe update:
- `TabularExport` now calls `sheet.FreezeRows(1)` after writing the
  header — the originally-specced "frozen header" behavior that was
  deferred when v0.3 introduced the IRow API.
- A new golden-file test asserts the freeze pane survives `Save →
  Open` round-trip.

Tests (+47 new, 252 per TFM total):
- `FreezeMergeHiddenTests` (17): freeze pane shape + round-trip;
  merge succeed / overlap-throws / adjacent-OK / 1×1 no-op /
  unmerge exact-match / unmerge non-match no-op / round-trip;
  bad-range rejection; sheet hidden round-trip; gridlines toggle;
  row hidden round-trip.
- `CellAddressTests`: +12 cases for `ParseRange` / `FormatRange`
  including the deferred-form rejections (`A:A`, `1:1`,
  `Sheet1!A1:B2`).
- Dispose-matrix: +13 new entries (8 ISheet, 2 IRow plus get/set
  variants).
- `TabularExportTests`: +1 case asserting the freeze landed.

### Diagnostic ID scheme unified (2026-05-16)
External review #N+1 flagged the dual diagnostic ID format: source-gen
diagnostics used `NXLS<NNNN>` (4-digit) while MSBuild build-time
guards used `NXLSAOT<NNN>` (category-prefix + 3-digit). Reviewer
warned this would compromise on the next category added; easier to
unify now than after publication.

- Renamed `NXLSAOT001` -> `NXLS0100` (PublishAot guard).
- Renamed `NXLSAOT002` -> `NXLS0101` (PublishTrimmed guard).
- Updated `buildTransitive/NetXlsx.targets`, README banner, design
  decision S27. Decision S16 (the ID-prefix decision) rewritten to
  document the range scheme:
  - `0001-0099` source-generator diagnostics
  - `0100-0199` MSBuild build-time guards
  - `0200-0299` reserved for Roslyn analyzers (v2+)
  - `0300+` reserved

No code logic changed; only the strings exposed to consumers. The
codes that previously shipped as `NXLSAOT001/2` only ever shipped
in v0.x preview packages, so there are no v1.0 consumers to break.

### v0.5 ReadRows slice — typed-mapping read path
The other half of `[Worksheet]` source-gen. `ReadRows` was the last
generator method still emitted behind `[Obsolete(error: true)]`;
that decoration is gone, and the body resolves headers + yields
records typed through the property map.

- **`Row_SheetExtensions.ReadRows(this ISheet, int? headerRow = 1)`**
  is now a real method. Body:
  1. Resolves the header row into a case-insensitive
     `Dictionary<string, int>` (matches design's culture rule).
  2. Looks up each `[Column(Name)]`-mapped property against the
     header map; throws `WorkbookException` if any header is missing.
  3. Iterates from `headerRow + 1` to the sheet's last row.
  4. For each row, checks if any mapped column has a value; skips
     fully-empty rows (continues — doesn't break, so an empty row
     in the middle is not the end-of-data marker).
  5. Yields `new T { ... }` with each property converted via the
     appropriate `GetX` cell-read + cast.
- **Conversion table** (per property type → cell-read expression):
  string → `GetString()`; bool → `GetBool() ?? throw`; numeric
  types → `GetNumber() ?? throw` with appropriate cast; DateTime
  → `GetDate() ?? throw`; DateOnly → `GetDateOnly() ?? throw`;
  TimeOnly → `GetTime() ?? throw`; TimeSpan → `GetDuration() ?? throw`.
  Required cells missing the expected value throw
  `WorkbookException` citing row + column-name + expected type.
- **Header-less mode deferred** (decision I-46): passing
  `headerRow: null` throws `NotSupportedException` with a "deferred
  to v2" message rather than silently doing the wrong thing.
- **Cookbook recipe 4 — `TypedImport`**. Round-trip recipe: writes a
  dataset via `TypedExport`'s path, reopens, reads back via the
  generated `ReadRows` extension. Golden-file test asserts the
  parsed records equal the input via `BeEquivalentTo`.
- Generator emission tests updated: `[Obsolete]` checks replaced with
  positive assertions ("ReadRows has a real body", "CS0619 no longer
  fires on calls"). The "no `[Obsolete]` in emitted output" assertion
  applies to the whole generated file now.

### v0.4.x small decisions batch (post-styling)
Three small concrete decisions from the design that hadn't been
implemented yet, plus the cookbook recipe each unblocks.

- **`CellError` enum + `ICell.GetError()`** (decision #49). Eight
  standard Excel error codes (#NULL!, #DIV/0!, #VALUE!, #REF!, #NAME?,
  #NUM!, #N/A, #GETTING_DATA). Maps NPOI's byte error codes to the
  typed enum. Returns null for non-error cells; surfaces error from
  formula cells with cached error results.
- **`Workbook.SuggestSheetName(IWorkbook, string)`** (design line 160).
  Returns the proposed name verbatim when unused; otherwise appends
  ` (2)`, ` (3)`, … until an unused name is found. Sanitizes invalid
  characters first. Truncates to 31 chars while preserving the
  disambiguating suffix. Case-insensitive collision detection
  (matches `AddSheet`'s duplicate rule).
- **Excel hard-limit enforcement on `SetString`** (decision #37 / §7.6).
  New `ResourceLimitExceededException` carries `LimitName` / `Limit` /
  `Actual`. Writes at exactly 32,767 chars succeed; one more throws.
  Hard limits for rows / columns are already enforced via
  `CellAddress.MaxRow` / `MaxColumn`; full `WorkbookOptions` for
  configurable limits is a follow-up.
- **Cookbook recipe 13 — `CellErrors`** (design §8.1). Writes seven
  Excel error codes — `#GETTING_DATA` is producible only by Excel's
  own evaluator from external data, not by NPOI's `SetCellErrorValue`
  — and demonstrates the read-side `GetError()` classification. The
  enum value remains in the API for files Excel authored.

Findings worth noting (captured for future implementation-notes):
- NPOI's `SetCellErrorValue` whitelists only seven of the eight codes;
  `#GETTING_DATA` (0x2B) throws `"Unknown error type: 43"`. Doesn't
  prevent reading it from real Excel files.
- The formula-cell-with-cached-error path through `GetError()` exists
  in code but requires the formula evaluator to materialize the cached
  state — out of scope for a unit test. Real-world Excel-authored
  workbooks exercise it.

### v0.4 styling slice
- **`Color` value type** (decision #29). Owned ARGB record struct, no
  `System.Drawing.Common` dependency. `FromRgb` / `FromArgb` / `FromHex` /
  `ToHex` + curated preset palette (Black, White, Red, Green, Blue,
  Yellow, LightGray, Gray). ARGB equality per decision I-23.
- **`CellStyle` value record** (design §6.8). Nullable per-axis properties:
  `Bold`, `Italic`, `Underline`, `FontName`, `FontSize`, `FontColor`,
  `Background`, `NumberFormat`, `HorizontalAlignment`, `VerticalAlignment`,
  `WrapText`, `Borders`. Null = "inherit existing on this axis." Structural
  equality drives style-pool dedup.
- **`CellBorders` record** + `BorderStyle` enum. Per-edge styles with
  optional per-edge colors; `CellBorders.All(style, color?)` helper.
- **`HAlign`, `VAlign`, `UnderlineStyle` enums**.
- **`NumberFormats` static class** (design §6.11, decision I12). Frozen
  v1.0 set: General, Text, Integer, Number, NumberTwo, Scientific,
  Percent, PercentTwo, Currency, CurrencyNoSymbol, Accounting, Date,
  DateTime, Time, Duration.
- **`ICell.Style(CellStyle)`** — merges over current style, resolves
  through the dedup pool, returns the cell for chaining.
- **`ICell.NumberFormat(string)`** — fluent shortcut for the common case.
- **`ICell.GetStyle()`** — returns the cell's current style as a
  `CellStyle` value record.
- **`CellStylePool` internal**. Per-workbook `Dictionary<CellStyle, ICellStyle>`
  keyed on `CellStyle` structural equality. Includes a separate font
  sub-pool keyed on `(name, size, bold, italic, underline, color)` so
  font-only differences don't allocate redundant NPOI `IFont` instances.
  **The S29 interim date/time style cache is gone** — the date-default
  styles are now regular pool entries; `SetDate`/`SetTime`/`SetDuration`
  flow through `StylePool.GetOrCreate(...)` like any other style.
- **Cookbook recipe 5 — `StyledReport`**. Bold + gray-filled centered
  header, currency-formatted Revenue column, yellow-highlighted rows
  for sub-15% margins. Demos all three primary axes (font, fill,
  number format) in one recipe. Golden-file test asserts the styles
  round-trip AND the dedup pool keeps the style index count small
  (proves the pool is actually deduping).

### Pre-styling cleanup (2026-05-16)
- `DisposedWorkbookMatrixTests` — parameterized matrix systematically
  verifying every public mutating member on `IWorkbook` / `ISheet` /
  `IRow` / `ICell` throws `ObjectDisposedException` after the owning
  workbook is disposed (decision #42). +55 cases. Adding a new public
  member is now a one-line `yield return` in the appropriate
  `MemberData`, not a copy-paste.
- `docs/scheduled-spikes.md` — quarterly cadence for re-checking NPOI's
  AOT/trim posture (spike 4-Q). Records past + future runs in a single
  table; documents promotion/demotion rules. Reviewer's recommendation
  #4 from the 2026-05-16 pass.

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
