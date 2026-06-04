# Changelog

All notable changes to NetXlsx are recorded here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html)
per `docs/design.md §3 #23`. Pre-1.0 minor versions may include breaking
changes (decision I19).

## [Unreleased]

### Added

- **Picture borders (I-86).** `IPicture.Border` — its first mutating member —
  reads and writes a solid line border (`<a:ln><a:solidFill>…`) on the
  picture's shape properties via the new `PictureBorder` record
  (`Color`/`ThemeColor`/`WidthPoints`; theme wins over explicit RGB per the
  I-79 precedence rule, theme aliases `tx1`/`bg1`/`tx2`/`bg2` are normalized
  on read). Set replaces the `<a:ln>` element wholesale and `null` removes
  it; get returns `null` for borders the record cannot represent faithfully
  (non-solid fills, unmapped scheme names, color-transform children) rather
  than approximating. Closes the ANIMAL PSS "10306 blister card" fidelity
  pin — a regenerated picture border now needs no escape-hatch authoring.

## [2.0.1] — 2026-06-04

### Fixed

- **The v2.0.0 bulk-write O(n²) regression (I-87).** The SDK engine's
  `GetOrCreateRow`/`MaxRowIndex` linearly scanned every `<row>` element on
  each call (~12 full scans per appended 10-cell row), making bulk in-memory
  writes quadratic — `Write5kRows` regressed from 251 ms (v1.x) to 3,652 ms
  and `StyledWrite_SmallPalette` from 8.8 ms to 41.4 ms, violating the
  design §4 targets (">500k styled cells/s", "30k rows < 3 s"). Row lookups
  are now served by a row-index cache with a documented escape-hatch
  coherence contract: accessing any `Underlying` member resets the caches
  (so acquire → mutate → continue-via-facade always observes hatch
  mutations), backstopped by per-access liveness checks. Within-row cell
  resolution gained the same treatment (tail-append fast path +
  last-resolved-cell memo), and repeat applications of the same
  `CellStyle` instance are served by a reference-keyed apply-memo (a memo
  hit still counts in the `StyleHitCount` diagnostic). Bulk writes return
  to the v1.x performance class and the design §4 budget (>500k styled
  cells/s; 30k rows < 3 s) is restored; no public API change.
  v2.0.0 was never published to NuGet; the first publishable version is
  ≥ v2.0.1.

### CI

- **The benchmark regression gate can actually fail now (I-87).** The
  previous gate compared against a rolling cache that refreshed on every
  main push with a soft-failing compare step — a regression that landed on
  main silently became the next run's baseline (how the O(n²) regression
  shipped with a green Benchmarks badge). The gate now compares every run
  against the committed `benchmarks/ci-baseline/` briefs, fails loud on a
  missing baseline, and is refreshed only as a deliberate, reviewable
  commit (see `benchmarks/README.md`).

## [2.0.0] — 2026-06-04

### v2.0.0 cutover — the Open XML SDK is THE engine (I-82) — BREAKING

The engine swap announced below is complete: every factory
(`Create` / `CreateMacroEnabled` / `CreateStreaming` / `Open` /
`OpenAsync`) now routes to the Open XML SDK engine, and NPOI is removed
from the library (it survives only as a test-side independent oracle).
`v2.0.0-alpha.1` is the release candidate cut from this state.

**Breaking — `.Underlying` retypes to the SDK objects.** Consumers
reaching through the escape hatch get a loud compile break and migrate
as follows:

| Member | Was (NPOI) | Now (Open XML SDK) |
|---|---|---|
| `IWorkbook.Underlying` | `XSSFWorkbook` | `DocumentFormat.OpenXml.Packaging.SpreadsheetDocument` |
| `ISheet.Underlying` | `XSSFSheet` | `DocumentFormat.OpenXml.Spreadsheet.Worksheet` (DOM root; part via `.WorksheetPart`) |
| `IRow.Underlying` | `XSSFRow` | `DocumentFormat.OpenXml.Spreadsheet.Row` |
| `ICell.Underlying` | `XSSFCell` | `DocumentFormat.OpenXml.Spreadsheet.Cell` (materializes the node on access — a write-like act) |
| `IPicture.Underlying` | `XSSFPicture` | `DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture` |
| `IShape.Underlying` | `XSSFSimpleShape` | `DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape` |
| `IConnector.Underlying` | `XSSFConnector` | `DocumentFormat.OpenXml.Drawing.Spreadsheet.ConnectionShape` |
| `IChart.Underlying` | `XSSFChart` | `DocumentFormat.OpenXml.Packaging.ChartPart` (DOM via `.ChartSpace`) |
| `ITable.Underlying` | `XSSFTable` | `DocumentFormat.OpenXml.Packaging.TableDefinitionPart` (DOM via `.Table`) |

**Breaking — streaming escape hatches removed.**
`IStreamingWorkbook.Underlying` and `IStreamingSheet.Underlying` are
gone, not retyped: the streaming engine writes rows forward-only
through `OpenXmlWriter` and assembles the package only at `Save`, so
there is no live document object to reach for at any earlier point —
a member that could never return would be a standing lie.

**Removed (never shipped).** `IWorkbook.OpenXmlDocument`, the swap-era
SDK hatch, is subsumed by the retyped `.Underlying`. The swap-era
factories `CreateOoxml` / `OpenOoxml` / `CreateStreamingOoxml` — exact
aliases of the default factories since the cutover — are gone too,
removed at the post-cutover fold slice (which also folded the
parallel-engine conformance project, `NetXlsx.OoxmlEngine.Tests`, into
`NetXlsx.Tests` under the `NetXlsx.Tests.Engine` namespace; no test was
lost). Note for `v2.0.0-alpha.1` users: that tag still carried the
aliases — replace them 1:1 with `Create` / `Open` / `CreateStreaming`.

**Unchanged on purpose.** `Open(Stream)` keeps its CanSeek +
Position==0 preconditions (relaxing to readable-only is additive and
deliberately deferred past v2.0.0). The public API surface is otherwise
identical — the facade is the asset; only the hatch types moved.

**Engine completions at the cutover.** `ICell.GetError` now reads
`t="e"` error literals on the SDK engine (plain and formula-cached,
decision #49 — including `#GETTING_DATA`, which NPOI's write API could
not produce); `CreateMacroEnabled` produces a genuine
`MacroEnabledWorkbook` (content type pinned by test).

**Dependency posture.** `DocumentFormat.OpenXml` is the library's only
runtime dependency. The NPOI package and the pre-v2 frozen-engine
security posture are gone (SECURITY.md rewritten); the
`SixLabors.ImageSharp` / `System.Security.Cryptography.Xml` transitive
pins remain only for the non-shipping test/bench projects that keep
NPOI as an oracle.

**AOT/trim unlocked.** The engine passed the `PublishTrimmed` +
`PublishAot` audit (zero IL/AOT warnings; representative workload
verified under a native binary), so the consumer-side `NXLS0100` /
`NXLS0101` build guards are removed. The library now also declares
`<IsAotCompatible>` (which implies `IsTrimmable`), so consumers' own
publish analyzers can see the claim; the trim/AOT/single-file
analyzers it enables run on every NetXlsx build and are clean on both
TFMs.

**Real-world gate.** The 5-file OPC stress sweep re-ran through the
now-default `Open` → `Save` on the flipped engine: 26/26, 31/31,
36/36, 17/17 and 50/50 parts preserved; every output reopens cleanly.

**Package metadata.** The NuGet `<Description>` and `<PackageTags>` no
longer describe the library as a facade over NPOI — the storefront now
matches the v2.0.0 reality (Open XML SDK engine, AOT/trim-compatible,
MIT all the way down). The repo README is now packed as the package
readme (`<PackageReadmeFile>`), so nuget.org renders a real landing
page instead of warning about a missing one.

**`FilterCriteria.In` limitation re-grounded (doc-only).** The 3+-value
`NotSupportedException` was documented as an NPOI 2.7.3 engine limit;
re-derived on the SDK engine it is model-shape-only (the engine fully
supports the `<filters>` value-list element — probe-verified under the
schema validator). The XML doc and design.md I-68 now say so, and full
value-list support is recorded as a queued post-v2.0.0 candidate. No
behavior change.

*(Cutover test migration: phase 1 moved the suite's NPOI reach-through
assertions to persisted-OOXML / public-API observables; phase 2
rewrote the deliberate residue against the SDK hatch and collapsed the
cross-engine differential harness — its de-risk mission completed at
the flip — onto literal-pinned single-engine contracts.)*

### Engine swap to Open XML SDK — foundation slice (I-82)

Begins the v2.0.0 engine swap from NPOI 2.7.3 to Microsoft's Open XML
SDK (`DocumentFormat.OpenXml` 3.5.1). The new engine grows **additively**
behind new factory methods — `Workbook.Create()`/`Open()` stay NPOI-backed
throughout the swap, so existing behavior and the full test suite are
unchanged.

- `Workbook.CreateOoxml()` / `Workbook.OpenOoxml(path)` /
  `Workbook.OpenOoxml(stream, leaveOpen)` — create/open on the Open XML
  SDK engine. This foundation slice supports create, open, save, dispose,
  `AddSheet`, sheet enumeration, and the indexers; other members throw
  `NotImplementedException` until their slice lands.
- `IWorkbook.OpenXmlDocument` — the SDK escape hatch
  (`SpreadsheetDocument?`). Returns the live document on the SDK engine,
  `null` on the legacy NPOI engine (the engine-discrimination signal).
  `IWorkbook.Underlying` (NPOI `XSSFWorkbook`) is unchanged on the legacy
  engine and throws `NotSupportedException` on the SDK engine.

**Cells & rows slice.** The SDK engine now reads and writes cell values:
`ICell.SetString` / `SetNumber` (double/decimal/int/long) / `SetBool` +
getters + `Kind` + `Clear`; `ISheet` indexers (`["A1"]`, `[r,c]`),
`Range(...)`, `AppendRow()` / `Row(int)`, `Column(...)`; `IRow` cell
navigation + `Set(...)` chaining; `IRange` sparse/dense enumeration,
`Value(object?)`, and `ClearContents`. Cells materialize lazily — reading
a never-written address reports `CellKind.Empty` without adding a node
(decision #40) — and rows/cells persist in Excel's required ascending
order. Date/time and formula setters, cell styling, comments, hyperlinks,
and rich text remain deferred to later slices (they throw
`NotImplementedException` on the SDK engine for now); a date in OOXML is a
number plus a date number-format style, so `SetDate`/`SetTime`/
`SetDuration` land with the styles slice.

**Cell styles slice.** The SDK engine now models `xl/styles.xml` and wires
the styling surface, backed by a new `OoxmlStylePool` that dedups
`CellStyle` values to a single `cellXfs` index (decision #4) and emits
OOXML schema types (`font`/`fill`/`border`/`numFmt`/`xf`) directly:

- `ICell.Style(CellStyle)` (merge semantics — non-null axes overlay, null
  inherit), `NumberFormat(string)`, `GetStyle()`, `ApplyNamedStyle(string)`;
  `IRange.Apply` / `ApplyNamedStyle` (dense); `IColumn.Width` / `WidthUnits`
  / `Hidden` / `SetDefaultStyle` / `ForEachPopulated`; `IRow.HeightInPoints`
  / `Hidden`; `IWorkbook.GetStylePoolDiagnostics` / `RegisterStyle` /
  `GetRegisteredStyle` / `RegisteredStyleNames`.
- The deferred date/time setters are unblocked: `ICell.SetDate` (DateTime /
  DateOnly) / `SetTime` / `SetDuration` write the Excel serial and apply the
  workbook's default date/time/duration number format when the cell is
  unstyled (decisions I-18/I-19, §7.9). `GetDate` / `GetDateOnly` and
  `Kind == CellKind.Date` detect date number formats (builtin ids 14–22 /
  45–47 plus custom date-token codes). The 1900/1904 epochs are honored,
  including Excel's fictitious 1900-02-29 leap day.
- `WorkbookOptions.DefaultFontName` / `DefaultFontSize` are applied to font
  index 0 on `CreateOoxml`; `DateSystem.Excel1904` is written to
  `workbookPr/@date1904`. On `OpenOoxml` the file's stylesheet and default
  font are adopted untouched (lessons #8, #9).

Deferred within the styles surface (land in later slices): `GetString` on a
date cell returns the raw serial rather than a format-rendered string (needs
a number-format renderer); named styles resolve in-memory but do not yet
persist into OOXML's `cellStyles` panel across save/open (the SDK equivalent
of the NPOI engine's I-67 round-trip); `IColumn.AutoSize` (needs font
metrics). Comments, hyperlinks, and formulas remain deferred.

**Rich text slice.** The SDK engine now reads and writes multi-run formatted
strings: `ICell.SetRichText(RichText)` / `GetRichText()`. A rich-text value
is written as an inline rich string (`<c t="inlineStr"><is>` with one `<r>`
per `RichTextRun`); each run's `<rPr>` carries its font axes inline
(bold/italic/underline/size/color/`rFont`), built by the style pool's
run-property helper. The marquee semantic (lesson #10): a run whose
`RichTextStyle` is empty is written with **no `<rPr>`**, so it inherits the
cell font — faithfully modeling the unformatted-prefix-run inheritance that
NPOI's fully-explicit run fonts could not preserve (the exact bug XlsxCodeGen
patched). A rich-text cell's `Kind` is `CellKind.String` and `GetString()`
returns the concatenated run text; `GetRichText()` returns `null` for a plain
string or non-string cell (it reads inline `<is>` runs and, on opened files,
shared-string `<si>` runs). Empty-text runs contribute no formatting run, and
the cell-text-length limit is enforced on the concatenated text.

**OPC preservation gate on the SDK engine (lesson #13).** Confirmed
structurally that an `OpenOoxml` → `Save` round-trip preserves every OPC part
NetXlsx does not model — the 70-85% file-size drop observed on real files is
pure SDK deflate recompression, not data loss (verified against all five
stress files: part sets preserved exactly, e.g. 50/50, 36/36, 17/17). A new
`OpcPreservationTests` builds a CI-safe fixture carrying an unmodeled custom
XML part and asserts the part set and the part's bytes survive Open → Save.

**Schema-validation gate (cross-cutting conformance).** The swap's founding
premise — that the Open XML SDK is schema-complete, so engine output is correct
"for free" — is now enforced rather than assumed. A reusable gate
(`OpenXmlValidationGate.AssertValid`) runs `DocumentFormat.OpenXml.Validation.
OpenXmlValidator` over a workbook's live `OpenXmlDocument` and asserts zero
errors, itemizing each error's part URI / element path / id / description on
failure. Fixtures cover every landed feature (scalar values; fonts/fills/borders/
numFmts/alignment; 1900- and 1904-epoch dates/times/durations; rich text incl. the
empty-style inheriting run) plus an `OpenOoxml`→`Save` round-trip; future slices
add their fixtures so the gate widens with the engine. Target:
`FileFormatVersions.Microsoft365` (also clean under `Office2019`). The gate found
the engine **already schema-clean** — including the rich-text `<rPr>` child order
that prior notes flagged as the prime suspect: the SDK does not constrain
`CT_RPrElt` child order (it does constrain `CT_Font` in `styles.xml`, and the font
path was already correct), so no engine change was required. (Operator-side
evidence: four of five real stress files round-trip schema-clean; the fifth carries
one source-authored `x14:workbookPr/@defaultImageDpi` value that the engine
OPC-preserves verbatim per lesson #13 — not engine-generated.) See
`docs/design.md` §6.2.15.

**Structure slice — merges + named ranges (first half).** The SDK engine now
implements the first half of the sheet/workbook structural surface (the slice is
split across sessions per its size):

- **Merges** — `ISheet.MergeCells` / `MergeCellsStyled` / `UnmergeCells` /
  `MergedRanges`. `<mergeCells>` is written after `<sheetData>` in the worksheet's
  strict child order; the container is dropped (not left childless) when the last
  region is removed. Behavior matches the NPOI engine: 1×1 merge no-op (I-38),
  overlap throws `InvalidOperationException` (§6.4), non-exact unmerge is a silent
  no-op, `MergedRanges` is canonical `A1:C3`, and `MergeCellsStyled` styles every
  cell in the range before merging so merged-region borders render from the
  boundary cells (lesson #4).
- **Named ranges** — `IWorkbook.AddNamedRange` / `NamedRanges` (+ `INamedRange`).
  Stored as `<definedNames>` in `workbook.xml`; sheet scope round-trips via
  `localSheetId`. Leading `=` is stripped, names are unique workbook-wide
  case-insensitively regardless of scope, and Excel name rules are validated
  (the SDK has no built-in validator). Contract matches the NPOI engine.

Both fixture families are added to the schema-validation gate (clean under
`Microsoft365`). No engine quirk surfaced. No public symbol added — every member
is an existing interface member newly implemented on the SDK engine, so the
PublicAPI snapshot is unchanged.

**Structure slice — panes / grouping / visibility / gridlines / width /
protection (second half).** Completes the sheet/workbook structural surface on the
SDK engine, each member writing the OOXML node directly and matching the NPOI
engine's contract (cross-checked against the NPOI-engine `FreezeMergeHiddenTests` /
`GroupingTests` / `SheetProtectionTests`):

- **Frozen / split panes** — `ISheet.FreezeRows` / `FreezeColumns` / `FreezePane`
  / `CreateSplitPane`. Written as `<sheetView><pane>` (xSplit = frozen columns,
  ySplit = frozen rows, `topLeftCell`, `activePane` + a matching `<selection>`,
  `state` = frozen or split). `(0, 0)` clears an existing freeze; negative args
  throw.
- **Grouping (outline)** — `GroupRows` / `UngroupRows` / `GroupColumns` /
  `UngroupColumns` / `SetRowGroupCollapsed` (1-based, validated). Row/column
  `outlineLevel` plus `<sheetFormatPr>` `outlineLevelRow`/`outlineLevelCol` tracking
  the deepest level; collapse hides the group block and marks the boundary row
  `collapsed`.
- **Visibility** — `ISheet.Hidden`, written as the `state` attribute on the
  workbook's `<sheet>` element (not the worksheet); both `hidden` and `veryHidden`
  read as hidden, clearing it on un-hide.
- **Gridlines / default column width** — `ShowGridlines` (`<sheetView>`
  `showGridLines`, default true, attribute cleared when set back to true) and
  `DefaultColumnWidth` (`<sheetFormatPr>` `defaultColWidth`; `null` clears it so
  Excel derives the width from font metrics, lesson #3 / I-78).
- **Sheet protection** — `Protect(password?, SheetProtection?)` / `Unprotect` /
  `IsProtected`. `<sheetProtection>` with `sheet=true`, all 15 granular lock flags
  set explicitly from the options, and the legacy 16-bit XOR verifier in
  `@password`; `Unprotect` removes the element.

A reusable schema-ordered insertion helper (`OoxmlSchemaOrder`) now places every
structural element at its correct position in the strict CT_Worksheet / CT_Workbook
child sequence. This fixes a latent ordering gap carried from the first half: the
bare `InsertAfter(<sheetData>)` / `InsertAfter(<sheets>)` inserts were correct only
for engine-*created* files, but on an *opened* file already carrying a legal
intervening sibling (`<autoFilter>` before `<mergeCells>`,
`<functionGroups>`/`<externalReferences>` before `<definedNames>`) they emitted
out-of-order XML that `OpenXmlValidator` rejects. The 5a merge / named-range inserts
are retrofitted onto the helper, and the gate gains open-mutate-validate fixtures
(inject a real-world intervening sibling, mutate, validate) so the blind spot is
closed for every structural slice from here (SDK-quirk #8).

**Schema-order completeness — machine-checked.** The helper's order lists were found
to be incomplete (a hand-maintenance hazard that had silently shipped gaps across two
slices). The lists are now derived from the SDK's *own* compiled schema particle and
locked by a guard test (`SchemaOrderCanonicalTests`): it reflects
`DocumentFormat.OpenXml`'s particle metadata for `CT_Worksheet`/`CT_Workbook` and
fails the build if either list drifts from the SDK by even one element. Ranking now
keys on element **local name** rather than CLR `Type` (the schema sequence is defined
over qualified names, and `<absPath>`'s class hides in an extension namespace a
`Type[]` would miss). Three previously-omitted elements are restored:
`<legacyDrawing>` / `<legacyDrawingHF>` (worksheet) and `<absPath>` (`x15ac:absPath`,
workbook ordinal 3 — emitted routinely by Excel and mis-ordered *today* by a shipping
`<definedNames>` insert before this fix). New open-mutate-validate fixtures cover both
the workbook `absPath` and worksheet `legacyDrawing` regions.

All structure-second-half fixtures are added to the schema-validation gate (clean
under `Microsoft365`). No public symbol added; the PublicAPI snapshot is unchanged.

**Drawings slice — pictures (first sub-slice).** The SDK engine now embeds and
reads back images. `ISheet.AddPicture` (all five overloads — single-cell, range,
each with explicit `ImageFormat` or magic-byte auto-detection, plus the range +
EMU-offset overload) and `ISheet.Pictures` are implemented on the SDK engine:

- Each sheet's images live in a `DrawingsPart` (`xl/drawings/drawingN.xml`,
  `xdr:wsDr`) referenced by a worksheet `<drawing r:id>` child, with the bytes in
  an `ImagePart` referenced by the picture's blip fill. The `<drawing>` child is
  placed by `OoxmlSchemaOrder`, so adding a picture to an *opened* sheet that
  already carries a later sibling (e.g. `<legacyDrawing>` for comments / form
  controls) keeps the worksheet schema-ordered (SDK-quirk #8).
- Single-cell overloads anchor at the cell's top-left at the image's **natural
  pixel size** via an `xdr:oneCellAnchor` with an EMU `<xdr:ext>` (PNG IHDR / JPEG
  SOFn dimensions × 9525 EMU/px) — no column-width-dependent resize. Range
  overloads anchor across two cells (`xdr:twoCellAnchor`), preserving per-image EMU
  offsets `dx1/dy1/dx2/dy2`, the end cell exclusive (lesson #5). The
  `FromCell`/`ToCell` read-back round-trips the 1-based addresses identically to the
  NPOI engine; a one-cell anchor reports `ToCell == FromCell` and `Dx2 == Dy2 == 0`.
- Magic-byte detection and validation mirror the NPOI engine exactly (PNG `89 50 4E
  47 …`, JPEG `FF D8 FF`; `UnsupportedImageFormatException` otherwise), de-risking
  the cutover. `IPicture.Underlying` (NPOI `XSSFPicture`) throws
  `NotSupportedException` on the SDK engine, the same escape-hatch divergence as
  `IWorkbook`/`ISheet`; the SDK package is reachable via `IWorkbook.OpenXmlDocument`.

Picture fixtures (both anchor kinds) and an open-mutate-validate fixture (add a
picture to an opened sheet carrying `<legacyDrawing>`) are added to the
schema-validation gate (clean under `Microsoft365`). Connectors / shapes and the
theme round-trip remain deferred to the rest of the drawings slice. No public symbol
added; the PublicAPI snapshot is unchanged.

**Drawings slice — shapes + connectors (second sub-slice).** `ISheet.AddShape`,
`ISheet.AddConnector`, and `ISheet.Connectors` are implemented on the SDK engine,
appending into the same `xdr:wsDr` root the pictures sub-slice builds (no new
part-type wiring; `xdr:wsDr` is not a strict-ordered container, so anchors append
freely):

- **Shapes** (`xdr:sp`) anchor across two cells with the end cell **exclusive** (the
  picture/shape convention). The public `ShapeType` maps to the OOXML preset geometry
  the NPOI engine emits — `rect` / `roundRect` / `ellipse` / `line` / `triangle` /
  `diamond`; `fillColor` → `<a:solidFill>` (else `<a:noFill>`), `lineColor` →
  `<a:ln><a:solidFill>`. `IShape` exposes only `Sheet`/`Type` (there is no
  `ISheet.Shapes` read-back), so this is a write-only fidelity surface.
- **Connectors** (`xdr:cxnSp`) anchor with the end cell **inclusive** (a connector's
  NPOI anchor maps the end cell with −1, so a same-cell connector round-trips
  `ToCell == FromCell`), preserving per-end EMU offsets. `ConnectorType` maps to the
  preset connector ordinals (`straightConnector1` = 96 / `bentConnector3` = 98 /
  `curvedConnector3` = 102, lesson #6); `lineColor`/`lineWidthPoints`/head+tail ends/
  `flipH`/`flipV` emit the matching `<a:ln w=…>` and `<a:xfrm>` flags. A fixed
  `<xdr:style>` block (lnRef idx 1 / `accent1`) matches the NPOI engine, so
  `LineStyleRefIndex` reads `1` and `LineSchemeColor` falls back to `"accent1"`.
  `ISheet.Connectors` reads them back with the full `IConnector` surface
  (`Type`/`From`/`ToCell`/`Dx1..Dy2`/`FlipH`/`FlipV`/`Head`+`TailEnd`/`LineColor`/
  `LineSchemeColor`/`LineWidthPoints`/`LineStyleRefIndex`).
- The geometry is asserted against the **NPOI engine as the parity oracle**
  (lesson #6: schema-valid ≠ positioned-correctly) — the SDK output emits the same
  preset ordinal, EMU markers, and `<a:ln>` line props the NPOI engine emits for the
  identical call, not merely valid XML. `IShape.Underlying` / `IConnector.Underlying`
  (NPOI types) throw `NotSupportedException`, the same escape-hatch divergence as
  pictures.

A shapes+connectors schema-valid fixture and a connector open-mutate-validate fixture
(add a connector to an opened sheet carrying `<legacyDrawing>`, exercising the shared
`<drawing>` schema-ordered insert) are added to the gate (clean under `Microsoft365`).
No public symbol added; the PublicAPI snapshot is unchanged.

**Drawings slice — theme round-trip (final sub-slice; completes the slice).**
`IWorkbook.SetThemeXml` / `GetThemeXml` and the read-side resolution
(`ResolveThemeColor` ×3 + `GetThemeLineWidthEmu`) are implemented on the SDK engine,
completing the drawings slice:

- The theme lives in `xl/theme/theme1.xml` as a dedicated `ThemePart` hung off the
  `WorkbookPart` by the theme relationship — **not** a child of any strict-ordered
  container, so the schema-order helper (SDK-quirk #8) does not apply here.
  `SetThemeXml` creates the part via `AddNewPart<ThemePart>()` (which also wires the
  content type + relationship) and writes the raw bytes with `FeedData`; a re-set
  reuses the existing part rather than orphaning a relationship. `GetThemeXml` reads
  the part stream directly (never materializing `ThemePart.Theme` into a DOM that
  would re-serialize and drift), so the bytes round-trip faithfully across save/open
  (OOXML truth — lesson #2: a missing theme breaks column-width display and
  theme-color resolution, so the theme must be preserved).
- The read side is engine-agnostic: it delegates to the shared `ThemeInfo`
  parsed-bytes view (decision I-81), so the OOXML cell-color slot mapping
  (`0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1..6, 10=hlink, 11=folHlink`), the
  `tx1`/`bg1`/`tx2`/`bg2` aliases, Excel's HLS tint algorithm, and the indexed
  line-width table behave identically to the NPOI engine. The cached `ThemeInfo` is
  invalidated by `SetThemeXml` so a re-set theme is resolved, not the stale one.
- The contract mirrors the NPOI-engine `ThemeReadAndDrawingIterationTests` (theme
  half) verbatim, de-risking the cutover; a theme schema-valid fixture is added to
  the gate (clean under `Microsoft365`). No public symbol added (all six are existing
  `IWorkbook` members newly implemented internally); the PublicAPI snapshot is
  unchanged. **This completes the drawings slice** — the next slice is conditional
  formatting / data validation / tables / autofilter / sort.

**Cross-engine differential harness + malformed-input fail-loud parity (I-83).**
The single biggest cutover de-risk: a `CrossEngineDifferentialTests` harness now
runs the **same scenario through both engines** (`Create()`/`Open()` = NPOI vs
`CreateOoxml()`/`OpenOoxml()` = SDK) and asserts they agree, across cell
values/kinds, 1900/1904 dates, rich-text runs, cell styles, merges, named ranges,
sheet visibility/gridlines, column width/hidden, row height/hidden, and two-cell
picture anchors. It compares semantically (not byte-identical XML), projecting out
the engines' legitimate unset-axis materialization differences (NPOI resolves an
unset number format to `"General"` / an unsized run to `FontSize 11`; the SDK
preserves the inherit semantic as `null`).

Its malformed-input half feeds hand-corrupted files through both engines and found
the SDK engine **silently defaulting** where NPOI **fails loud** — a corrupt
shared-string `<v>` index read back `""`, an unparseable numeric `<v>` read back
`0`, a non-integer drawing-anchor marker read back as column 0 (mis-placing the
drawing). **Decision I-83 aligns the SDK engine to fail loud** with
`MalformedFileException` at these sites (shared-string index resolution, numeric
parse, anchor markers — all in `Internal/Ooxml*.cs`), restoring the library's
fail-loud honesty contract. OOXML defines no default for any of these corrupt
values. A corrupt boolean value was reviewed and deliberately left lenient (both
engines already read `false`). No public symbol added; no behavior change on
well-formed files.

**CF / data validation / tables / autofilter / sort slice.** The SDK engine
now carries the full structured-data surface, each feature oracle-checked
against the NPOI engine's emitted XML and covered by the differential harness:

- `ISheet.SortRange` (I-72): stable in-memory sort with NPOI's exact
  comparison semantics (blanks last, numbers before strings,
  ordinal-ignore-case, FALSE < TRUE). The engine detaches and re-homes the
  in-range `<c>` elements, so styles, inline rich text, and formula text move
  verbatim by construction (including the documented formula caveat —
  references are NOT relocated).
- AutoFilter (I-56/I-66): `SetAutoFilter` / `ClearAutoFilter` / per-column
  `FilterCriteria` / `HasAutoFilter` / `AutoFilterRange`. Also
  creates/updates Excel's hidden `_xlnm._FilterDatabase` built-in name
  (NPOI parity, oracle-pinned: quoted-when-needed sheet name, no 1×1
  collapse, name survives `ClearAutoFilter`, filter columns survive a range
  re-set).
- Data validation (I-55): every `DataValidation` factory family. The type
  gains an engine-agnostic internal descriptor (the OOXML
  type/operator/formula axes) alongside its NPOI-typed closure — no public
  change; the explicit-list quoted-join encoding matches NPOI exactly.
- Conditional formatting (I-73): cellIs / expression / colorScale rules with
  dxf (differential-format) styles modeled in the style pool (deduped,
  unlike NPOI's one-per-rule). Documented divergences from NPOI cosmetics:
  schema-default attributes omitted, conformant 8-digit ARGB colors, no
  meaningless dxfId on colorScale rules, and rule priorities allocated
  max+1 so a remove-then-add can never mint NPOI's duplicate priorities.
- Excel tables (I-51/I-64): `AddTable` / `Tables` / `TryGetTable` /
  `RemoveTable` + the full `ITable` surface incl. the totals-row lifecycle
  (SUBTOTAL 100-series codes, custom formulas, labels). The engine EMITS the
  schema-required `<table @id>` that NPOI 2.7.3 omits. `ICell.GetFormula`
  now reads formula cells back (`"="` + body, NPOI parity) — formula cells
  are producible on this engine via opened files and the totals writer.

`<conditionalFormatting>` is the first 0..* member of CT_Worksheet's strict
child sequence the engine writes; `OoxmlSchemaOrder` gains an `Insert`
companion that always creates + positions (repeats keep document order),
alongside the existing `GetOrInsert` for 0..1 singletons. Every new
strict-ordered element has an open-mutate-validate schema fixture
(SDK-quirk #8), and the new opened-file parse site (autoFilter `@ref`) is
checked by the malformed-input harness (SDK-quirk #13).

**Fixed: conditional-format Bold/Italic swap on the NPOI engine.** NPOI's
`IFontFormatting.SetFontStyle` takes `(italic, bold)` — POI's parameter
order — but `ConditionalFormat.ApplyStyle` passed `(bold, italic)`, so a CF
rule styled Bold rendered Italic in Excel and vice versa (invisible when
both flags were set). Found by the new cross-engine emission-parity test,
which asserts the two engines' emitted cfRule + dxf shapes agree.

**Charts slice.** The SDK engine now carries `ISheet.AddChart` (I-75) for
all six `ChartType`s — line, bar, column, pie, scatter, area — plus
`IChart.SetTitle`, oracle-checked against the NPOI engine's emitted chart
XML and pinned by a cross-engine emission-parity test:

- A chart lands as a `ChartPart` referenced from an `xdr:graphicFrame` in a
  two-cell anchor (end cell exclusive, the picture/shape convention). The
  series caches (`strCache`/`numCache`) snapshot the referenced cells at
  `AddChart` time with NPOI's exact skip semantics: `ptCount` covers the
  full range; only type-matching cells emit a point (date cells count as
  numeric serials). References are sheet-qualified absolute, quoted when
  needed, single-cell ranges collapsed.
- `IChart.Underlying` (NPOI `XSSFChart`) throws `NotSupportedException` on
  the SDK engine — the same escape-hatch divergence as
  `IPicture`/`IShape`/`IConnector`/`ITable`.
- Documented divergences from NPOI cosmetics/nonconformities (the schema
  gate wins): the pie `dPt` accent list is emitted in CT_PieSer schema
  order; no dangling axes on pie charts; scatter charts plot on two value
  axes per ECMA-376 (NPOI pairs a category axis); drawing `cNvPr` ids are
  unique-nonzero (NPOI emits the invalid `id="0"`); the chart part lives at
  `xl/drawings/charts/chartN.xml` (relationship-resolved; NPOI/Excel use
  `xl/charts/`).

**Formulas / comments / hyperlinks + workbook protection slice.** Closes
every remaining `ICell` stub and the non-streaming `IWorkbook` stubs on the
SDK engine — only `IColumn.AutoSize` (font metrics) and the streaming
surface remain:

- `ICell.SetFormula` writes the bare `<c><f>` shape with no cached value
  (design §7.8 / #46). The SDK has no formula parser (NPOI does), so the
  engine validates *structure* — balanced parentheses and terminated
  string/quoted-sheet-name literals — failing loud with `FormulaException`
  on the breakage Excel would reject; semantic validity stays Excel's call.
  Documented divergence under I-82.
- `ICell.Comment`/`GetComment`/`GetCommentAuthor` build the full two-part
  comment graph: the comments part (authors deduped, text preserved) plus
  the VML legacy-drawing popup shape wired by a schema-ordered
  `<legacyDrawing r:id>` — geometry matching NPOI exactly, collision-safe
  shape ids on opened files, mutate-in-place replace semantics.
- `ICell.Hyperlink`/`GetHyperlink` reproduce the I13 scheme sniff:
  http(s)/mailto/file as external package relationships, `#Sheet!Range` as
  `@location`, everything else rejected. Targets round-trip VERBATIM (no
  URI canonicalization); replacing a link removes the superseded
  relationship.
- `IWorkbook.Protect`/`ProtectWithPassword`/`Unprotect`/`IsProtected`
  (I-54/I-65) write the schema-ordered `<workbookProtection>` with explicit
  flags and the legacy 16-bit XOR verifier (byte-identical to NPOI's, so
  passwords stay cross-engine compatible); `IsMacroEnabled` reads the
  document content type, detecting NPOI-authored `.xlsm` files.
- New fail-loud parse sites per I-83: a dangling hyperlink `r:id` and an
  out-of-range comment `authorId` throw `MalformedFileException`; a
  non-integer `authorId` fails loud too (NPOI silently coerces it to the
  empty author — the silent default I-83 forbids; deliberately stricter).
- The O-9 strict-floor tripwire: a kitchen-sink workbook spanning every
  implemented surface now validates under `Office2007` in CI, so the first
  unwrapped post-2007 extension construct forces a conscious
  wrap-or-raise-the-floor decision. The primary schema gate remains
  `Microsoft365`.

**Streaming write slice (slice 9 — the SXSSF replacement).** The last big
surface on the SDK engine: the `IStreamingWorkbook` / `IStreamingSheet` /
`IStreamingRow` / `IStreamingCell` contract (decision #7 / I-13), reached
through the new additive factory:

- `Workbook.CreateStreamingOoxml(StreamingOptions? options = null)` — the
  Open XML SDK counterpart to `CreateStreaming`, which stays NPOI-backed
  until cutover. **No `IStreaming*` interface change was needed**: the I-13
  surface is already append-only across rows, and within-window cell
  mutation maps directly onto the `RowAccessWindowSize` buffering SXSSF
  itself performs — the forward-only `OpenXmlWriter` shape is honored, not
  papered over.
- Each sheet streams its worksheet XML into its own temp file through an
  `OpenXmlPartWriter` as rows leave the bounded window; `Save` assembles the
  package by feeding the finished worksheet bytes into the parts — memory
  stays bounded end-to-end. `RowAccessWindowSize`, `CompressTempFiles`
  (gzip temp files), `DateSystem`, and the default-font options are all
  honored.
- Cell emission matches the random-access SDK engine: inline strings with
  space preservation, G17 invariant numbers, no cached `<v>` under formulas
  (#46), `SetDate` = serial + default datetime format when unstyled. Styles
  dedup through the same pool semantics (#4/#29) via a detached stylesheet
  attached at Save.
- Fail-loud contract (documented divergences from NPOI's SXSSF, probed
  before implementing): `Save` is single-shot on both engines, but NPOI
  leaks `ObjectDisposedException` from writer internals on a second save
  and silently DISCARDS rows appended after save or writes to
  window-evicted rows — the SDK engine throws a deliberate
  `InvalidOperationException` at all three sites (I-83 honesty). A failed
  save to a bad path does not burn the single shot. The NPOI escape hatches
  (`Underlying`) throw `NotSupportedException` on this engine.
- The two slice oracles: a cross-engine streaming differential (same
  dataset streamed through both engines, both outputs reopened through both
  readers, all four projections equal) and a forward-only/bounded-memory
  guard (a window-evicted row rejects writes — re-introducing hidden
  buffering breaks the test). Streamed output is schema-valid under
  `Microsoft365`.

**Closeout slice (date-cell `GetString` rendering + named-style
persistence).** Empties the within-surface deferral list except
`IColumn.AutoSize`:

- A date-formatted numeric cell's `GetString` now renders through its
  number format (design §7.10) instead of returning the raw serial — new
  internal `ExcelDateFormat` renderer covering y/m/d/h/s fields with
  Excel's month-vs-minute context rule, elapsed `[h]`/`[m]`/`[s]` tokens,
  `AM/PM` + `A/P` meridiems, millisecond fractions, quoted/escaped
  literals, sections, and color/locale prefixes. Oracle-pinned against
  NPOI's `DataFormatter` over a 40+ format matrix; output is
  culture-independent (matching NPOI, verified under de-DE). Negative
  serials fall back to raw G17 on both engines. Where NPOI demonstrably
  mangles (quoted literals, lowercase meridiems, `A/P`,
  meridiem-before-hour), the SDK engine renders Excel-correct — each
  divergence pinned with NPOI's actual output noted inline.
- `RegisterStyle` on the SDK engine now persists named styles into the
  stylesheet's `cellStyleXfs`/`cellStyles` tables (the NPOI engine's I-67
  round-trip): names survive save/open, rehydrate on first access, and
  appear in Excel's Cell Styles panel. User entries carry no `builtinId`
  (NPOI stamps `builtinId="0"`, which claims the Normal builtin per
  ECMA-376 — a witness artifact, not a contract); cross-engine
  rehydration is pinned in both directions. A stylesheet missing the
  style tables entirely gets them created in schema order with the
  Normal master seeded at index 0, keeping existing `xfId="0"`
  references Normal-shaped.

**AutoSize slice (embedded font metrics — I-84).** `IColumn.AutoSize` lands
on the SDK engine, emptying the stub inventory: **the SDK engine is now
functionally complete** — every `IWorkbook`/`ISheet`/`IRow`/`ICell`/`IRange`/
`IColumn` member is implemented.

- AutoSize measures with embedded numeric font-metric tables
  (`Internal/FontMetricsTables.g.cs`, generated by the new
  `tools/FontMetricsGen` from openly-licensed metric-compatible fonts —
  Carlito ↔ Calibri/Aptos, Liberation Sans ↔ Arial/Arimo/Helvetica,
  Liberation Serif ↔ Times New Roman/Tinos, Liberation Mono ↔ Courier
  New/Cousine; only numbers ship, never font bytes, so the MIT posture is
  untouched). Zero runtime dependency; deterministic on every machine,
  headless included; AOT-clean. Width math mirrors NPOI 2.7.3's
  `SheetUtil`, oracle-dumped before implementation — merged-region cells
  skipped, per-line measurement, all-empty columns are a no-op, 255-unit
  cap, `bestFit`+`customWidth` on the emitted `<col>`.
- Fonts outside the embedded set throw `MissingFontException` naming the
  font (decision I-84, superseding I3's environment-dependence on this
  engine: the failure mode is now "font not in the embedded metric set",
  identical everywhere, instead of "host has no fonts").
- Documented divergences from the NPOI engine (design.md I-84): advance-sum
  measurement (~1px wide of NPOI's ink box), non-date custom formats
  measure shortest round-trip text, fresh formulas with no cached result
  are skipped, and cross-engine width parity is tolerance-based (25%),
  not exact — NPOI's widths are environment-dependent by construction.

**Pre-cutover parity slice (closes the O-15 scout's (a)-list).** The
cutover-readiness scout (2026-06-03, throwaway branch) ran the full v1.3
suites against the SDK engine and surfaced 8 real divergences beyond the
expected `.Underlying` reach-throughs; this slice closes all of them:

- **WorkbookOptions read limits** (`ReadMaxSheets` /
  `ReadMaxUncompressedBytes`) are enforced post-open on `OpenOoxml`, and
  the **write caps** (`MaxRowsPerSheet` / `MaxColsPerSheet`) gate the sheet
  indexer, `Range` corners, `AppendRow`, `Row(int)`, `Column(int)` and
  `IRow.Cell(int)` with the NPOI engine's exception wording.
- **Concurrency contract**: decision #43's opportunistic reentry counter
  (raced structural mutation → `InvalidOperationException` naming
  `StrictConcurrencyDetection`) and decision I-59's strict real-lock mode
  now guard `AddSheet` / `AddNamedRange` on the SDK engine. Previously a
  raced `AddSheet` corrupted the package.
- **`IRange` parity**: the bounds properties gained disposed-workbook
  guards (decision #42), and `IRange.Value` now dispatches the unsigned
  scalars (`ushort`/`uint`/`ulong`) and throws the pinned
  "is not a supported scalar" message.
- **Relationship-orphan OPC parts survive Save** (decision #44): the SDK
  part graph is relationship-defined, so legal-OPC zip entries with no
  .rels chain were silently dropped by the clone-based Save. The engine
  now captures orphans at Open and re-injects them into every Save,
  byte-identical.
- **Open malformed-input gate** (the I-60 equivalent): empty input, zips
  with no workbook part, and lazily-surfaced package corruption now throw
  `MalformedFileException` from `OpenOoxml` instead of opening a phantom
  package or leaking raw `InvalidOperationException`.

### Added — `ISheet.LastRowNumber` (I-85)

`int ISheet.LastRowNumber { get; }` — the 1-based index of the last row
containing at least one cell, `0` when the sheet has no cells. An
empty-but-materialized row (from `Row(int)` / `AppendRow()`) does not
count; a cleared cell still does. Implemented on both engines with
identical semantics.

Motivation: the source generator's emitted `ReadRows<T>` body reached
through `sheet.Underlying.LastRowNum` (NPOI-typed, 0-based) — generated
consumer code that would break at cutover. The generator now emits
`sheet.LastRowNumber`; regenerated code is engine-agnostic. (A
generator-side row-probing heuristic was rejected — it would bake an
approximation into every consumer's compiled assembly.)

No breaking change in these slices. The `.Underlying` return-type change,
the NPOI removal, and the default-engine cutover land together in a later,
focused **v2.0.0** cutover slice, gated on the full suite passing against
the SDK engine. See `docs/design.md` I-82.

Coverage: `tests/NetXlsx.OoxmlEngine.Tests/` (`FoundationRoundTripTests`,
`CellAndRowValueTests`, `CellStyleTests`, `RichTextTests`,
`OpcPreservationTests`, `SchemaValidationTests`, `SchemaOrderCanonicalTests`,
`MergeTests`, `NamedRangeTests`, `PaneTests`, `GroupingTests`,
`SheetStructureTests`, `SheetProtectionTests`, `PictureTests`,
`ShapeConnectorTests`, `ThemeTests`, `SortTests`, `AutoFilterTests`,
`DataValidationTests`, `ConditionalFormatTests`, `TableApiTests`,
`ChartTests`, `FormulaTests`, `CommentTests`, `HyperlinkTests`,
`WorkbookProtectionTests`, `CrossEngineDifferentialTests`,
`CrossEngineMalformedInputTests`, `StreamingEngineTests`,
`DateGetStringTests`, `NamedStylePersistenceTests`, `AutoSizeTests`,
`WorkbookOptionsParityTests`, `ConcurrencyContractTests`,
`RangeContractParityTests`, `OpenMalformedGateTests`,
`OrphanPartPreservationTests`, `LastRowNumberTests`).

### Read-side introspection: themes + drawings (I-81)

Adds the symmetric read surface that consumers (reproduction tools,
validators, converters) were previously building by reaching through
`.Underlying`:

- `IWorkbook.GetThemeXml()` (counterpart to `SetThemeXml`),
  `ResolveThemeColor` (three overloads: by index, by `ThemeColor`, by
  scheme name with `tx1`/`bg1` aliases) using Excel's correct tint
  algorithm, and `GetThemeLineWidthEmu(int)`.
- `ISheet.Pictures` / `ISheet.Connectors` — drawing-order read-only lists.
- `IPicture` gains `FromCell`/`ToCell`/`Dx1..Dy2`/`Data`.
- `IConnector` gains anchor + offsets, `FlipH`/`FlipV`, `HeadEnd`/`TailEnd`,
  `LineColor`, `LineSchemeColor`, `LineWidthPoints`, `LineStyleRefIndex`.

Coverage: `tests/NetXlsx.Tests/ThemeReadAndDrawingIterationTests.cs` plus
new `DisposedWorkbookMatrix` rows.

### Source generator: inheritance, incrementality, honest NXLS0006

- `[Worksheet]` types now map public properties inherited from base
  classes (base columns first), instead of silently dropping them.
- `PropertyLocation` (file/line/column) is excluded from the cached
  model's equality, so trivia-only edits no longer bust generator
  incrementality.
- NXLS0006's message no longer advertises `Nullable<T>` as a built-in
  supported type (the generator rejects it); it now points to the custom
  converter (`[Column(ConverterType = ...)]`) for unmodeled types.

Coverage: new `Inherited_Public_Columns_Are_Mapped_Base_First` emission
test.

### Fix: wire `WorkbookOptions.DateSystem` (1904 epoch)

`WorkbookOptions.DateSystem` was declared and defaulted but read nowhere,
so `DateSystem.Excel1904` silently did nothing — a spec-vs-code
divergence against design decision #15. New workbooks created with
`Excel1904` now set the `workbookPr/@date1904` flag, so dates serialize
and read back against the 1904 epoch (random-access and streaming).
Opening an existing file never overrides the file's own epoch.

Coverage: `tests/NetXlsx.Tests/DateSystemTests.cs` (7 tests).

### Connector enhancements (I-80)

Reworks the I-79 connector API for faithful reproduction of arrows and
lines:

- `ConnectorType` values corrected to `ST_ShapeType` ordinals
  (`Straight = 96`, `Bent = 98`, `Curved = 102`). The I-79 values mapped
  `Straight` to `star8`, so every connector rendered as a star.
- `ISheet.AddConnector` now accepts EMU offsets (`dx1..dy2`),
  `flipH`/`flipV`, head/tail arrowheads (`ConnectorEnd`), and
  `lineWidthPoints`, and returns the new `IConnector` facade instead of a
  raw `XSSFConnector`.
- New `ConnectorEnd` enum and `IConnector` interface.

Coverage: `tests/NetXlsx.Tests/ConnectorTests.cs` plus a new
`DisposedWorkbookMatrixTests` row.

### `.xlsm` macro-enabled passthrough (I-69)

Adds `Workbook.CreateMacroEnabled(options?)` to create macro-enabled
workbooks and `IWorkbook.IsMacroEnabled` to detect them. NPOI 2.7.3's
`XSSFWorkbook(XSSFWorkbookType.XLSM)` constructor handles the OOXML
content-type switch cleanly; `Workbook.Open` already handles `.xlsm`
files transparently (NPOI detects the content type).

- New `Workbook.CreateMacroEnabled(WorkbookOptions? options = null)`.
- New `IWorkbook.IsMacroEnabled` read-only property.
- `Workbook.Open` doc-comments updated to mention `.xlsm`.
- NetXlsx does not read, write, or execute VBA — macro content is
  passthrough only (VBA project parts survive round-trip via NPOI's
  OPC-part preservation, decision #44).

Coverage: 8 new tests in `tests/NetXlsx.Tests/XlsmPassthroughTests.cs`
— create, open, stream round-trip, file round-trip, double round-trip,
options forwarding, regular `.xlsx` negative check.

Decision added: I-69.

### Split panes (I-70)

Adds `ISheet.CreateSplitPane(xSplitTwips, ySplitTwips)` — a
draggable (non-frozen) pane split. Parameters are in twips (1/20th
of a point), matching NPOI and OOXML's coordinate system. Replaces
any prior freeze or split on the sheet.

Coverage: 7 new tests in `tests/NetXlsx.Tests/SplitPaneTests.cs`.

Decision added: I-70.

### Grouping / outlining (I-71)

Adds row and column grouping (Excel's outline feature):
- `ISheet.GroupRows(startRow, endRow)` / `UngroupRows`
- `ISheet.GroupColumns(startCol, endCol)` / `UngroupColumns`
- `ISheet.SetRowGroupCollapsed(row, collapsed)`

Nested groups supported (NPOI increments outline level per nesting).
All indices are 1-based per NetXlsx convention.

Coverage: 12 new tests in `tests/NetXlsx.Tests/GroupingTests.cs`.

Decision added: I-71.

### Sorting helpers (I-72)

Adds `ISheet.SortRange(a1Range, params SortKey[] keys)` — pure facade
that physically reorders cell values and styles within a range. New
`SortKey` type with `Asc(col)` / `Desc(col)` factories (1-based
columns). Sort order matches Excel: numbers before strings, blanks
last, case-insensitive string comparison. Multiple keys applied in
declared order. No OOXML sortState metadata written.

Coverage: 12 new tests in `tests/NetXlsx.Tests/SortTests.cs`.

Decision added: I-72.

### Conditional formatting (I-73)

Adds `ISheet.AddConditionalFormatting(a1Range, params ConditionalFormat[]
rules)` with factories for cell-value conditions (7 operators),
formula-based rules, and 2/3-color gradient scales. Also
`ConditionalFormattingCount` and `RemoveConditionalFormatting(index)`.

Style application supports bold/italic font formatting and fill
background color via XSSF's RGB path. Font color in CF rules deferred
(NPOI 2.7.3 limitation). Data bars, icon sets, top/bottom N reach
through `ISheet.Underlying.SheetConditionalFormatting`.

Coverage: 15 new tests in `tests/NetXlsx.Tests/ConditionalFormatTests.cs`.

Decision added: I-73.

### Drawings / shapes (I-74)

Adds `ISheet.AddShape(type, startCell, endCell, fillColor?, lineColor?)`
with `ShapeType` enum (Rectangle, RoundedRectangle, Ellipse, Line,
Triangle, Diamond) and `IShape` interface. Shape anchors span between
two cells. Exotic shapes and advanced properties (rotation, text,
gradients) reach through `IShape.Underlying` or `ISheet.Underlying`.

Coverage: 8 new tests in `tests/NetXlsx.Tests/ShapeTests.cs`.

Decision added: I-74.

### Charts (I-75)

Adds `ISheet.AddChart(type, startCell, endCell, categoryRange,
valueRange, title?)` with `ChartType` enum (Line, Bar, Column, Pie,
Scatter, Area) and `IChart` interface. Single-series charts backed
by NPOI's `IChartDataFactory`. Multi-series and advanced customization
reach through `IChart.Underlying`.

Coverage: 10 new tests in `tests/NetXlsx.Tests/ChartTests.cs`.

Decision added: I-75.

### Excel-correct defaults + DefaultColumnWidth (I-78)

Fixes two NPOI 2.7.3 defaults that cause every consumer's workbooks
to display incorrect column widths in Excel:

- **Normal cellStyle entry:** `Workbook.Create()` now ensures
  `<cellStyle name="Normal" builtinId="0"/>` exists in styles.xml.
  Without it, Excel cannot resolve the Normal style → font → Maximum
  Digit Width chain for default column width calculation.
- **Suppress `defaultColWidth`:** new sheets no longer emit
  `defaultColWidth="8.43"` in `<sheetFormatPr>`. When present, Excel
  uses the literal float (which rounds to 7.71 display). When absent,
  Excel derives the width from font metrics (correct 8.43 display).
- **`ISheet.DefaultColumnWidth`:** new nullable double property.
  `null` (default) omits the attribute; non-null writes it explicitly.

Coverage: 5 new tests in `DefaultColumnWidthTests.cs`.

Decision added: I-78.

### Row height + two-cell image anchoring (I-76)

Closes two gaps identified when analyzing a form-layout xlsx file:

- `IRow.HeightInPoints` — get/set row height in points. Enables
  precise vertical layout for form-style sheets with varying row
  heights (e.g. 5.25pt spacers, 24.95pt data rows, 80pt description
  rows).
- `ISheet.AddPicture(startCell, endCell, data, format)` — two-cell
  anchor overload that stretches/shrinks the image to fill the anchor
  region. Unlike the single-cell overload (I-52) which renders at
  natural pixel size, this gives layout control for positioned images.

Coverage: 7 new tests in
`tests/NetXlsx.Tests/RowHeightAndPictureAnchorTests.cs`.

Decision added: I-76.

## [1.3.0] — 2026-05-22

v1.3 ships two slices that close the remaining v1.2 deferments
plus an upstream-gap acknowledgement for the third item.

- **Slice 1 (I-67):** OOXML named-style table integration.
  `IWorkbook.RegisterStyle` now writes to `cellStyleXfs` +
  `cellStyles`; `RegisteredStyleNames` / `GetRegisteredStyle`
  lazy-rehydrate from OOXML on first access. Named styles
  survive `Workbook.Open` round-trip and appear in Excel's
  Cell Styles ribbon.
- **Slice 2 (I-68):** `FilterCriteria.In(values)` — partial
  landing. 1–2-value support via the customFilters route;
  `NotSupportedException` for 3+ values (NPOI 2.7.3's
  `CT_FilterColumn` doesn't surface the `<filters>` element).
- **Slice 3 (deferred indefinitely):** `FilterCriteria.Top(n)` /
  `BottomPercent(...)`. Unlike `In(...)`, Top-N has no
  customFilters fallback — adding a method that always throws
  would be a footgun. Held in the roadmap pending an NPOI 3.x
  bump that surfaces `top10`.

Decisions added: I-67, I-68.

Test totals: **612 unit + 35 golden-file + 1 public-API
snapshot + 18 fuzz = 666/TFM × 2 TFMs = 1,332 total** runs per
CI build (was 655/TFM at v1.2). PublicAPI.Shipped.txt = 571
entries (was 569 at v1.2; +2 for the `In(...)` overloads —
v1.3's other behavioral changes don't add public surface).

### v1.3 slice 2 — `FilterCriteria.In(...)` (partial, I-68)

Adds the explicit-value-list filter factory deferred from v1.2.
NPOI 2.7.3's `CT_FilterColumn` proxy models only `customFilters`,
not the OOXML `<filters>` element — so v1.3 ships **two-value
support via the customFilters route** (the same path Excel uses
internally for short value lists) and throws
`NotSupportedException` for 3+ values.

- New `FilterCriteria.In(params string[] values)`.
- New `FilterCriteria.In(IEnumerable<string> values)`.
- 0 values → `ArgumentException`.
- 1 value → reduces to `EqualTo(v[0])`.
- 2 values → composes `EqualTo(a).Or(EqualTo(b))`.
- 3+ values → `NotSupportedException` with a message naming NPOI
  2.7.3 and the `<filters>` element. Lifts cleanly when an NPOI
  3.x bump or a future XML-emission workaround removes the limit
  — no caller-code change required.

Coverage: 7 new tests in
`tests/NetXlsx.Tests/AutoFilterTests.cs` — single-value reduces,
two-value OR-joined, three-or-more throws, empty rejected, null
entries rejected, null argument rejected, IEnumerable overload
exercised.

### v1.3 slice 1 — OOXML named-style table integration (I-67)

Closes the slice deferred from v1.2 (originally roadmap line item
under v1.2 §"OOXML named-style table integration", reassigned to
v1.3 after in-flight scope assessment).

- **`IWorkbook.RegisterStyle(name, style)` now writes to OOXML.**
  Each call adds a CT_Xf to the `cellStyleXfs` table and a
  CT_CellStyle entry (name + xfId) to `cellStyles`. The
  in-process name → CellStyle map remains the runtime source of
  truth; OOXML serialization is a side effect.
- **`IWorkbook.RegisteredStyleNames` and `GetRegisteredStyle` now
  rehydrate from OOXML on first access** after a workbook is
  opened. A freshly-opened workbook surfaces all named styles
  that the saving tool wrote to `cellStyles`, except the built-in
  "Normal" entry (Excel always creates one; we suppress it to
  keep `RegisteredStyleNames` user-meaningful).
- Both write and read sides use `Internal/NpoiInternals.cs` for
  the reflection-bounded NPOI 2.7.3 internals access
  (`StylesTable.PutCellStyleXf`, `GetCellStyleXfAt`, `PutCellXf`).
  Each access is one cached `MethodInfo` at class-init; an NPOI
  version that moves any of these throws a clear "internal
  change" diagnostic at first-use.
- **Named-style replacement** (`RegisterStyle("X", ...)` twice)
  updates both the in-process map and the OOXML entry in place
  — does not append a duplicate `<cellStyle>` row.
- **Visual style preservation across round-trip:** the read path
  materializes each named-style XF as a regular cellXfs entry and
  pipes through `CellStylePool.ReadFromNpoi(ICellStyle)` to
  produce a `CellStyle` value. One duplicate cellXfs entry per
  named style at read time; cost is bounded.
- **Carried-forward limitation:** cells styled via
  `ApplyNamedStyle` do not carry a `xfId` reference to the
  named-style XF after round-trip; they get an explicit cellXfs
  entry. Visual outcome is identical; Excel's Cell Styles ribbon
  group won't highlight the cell's style by name. Closing that
  gap is a future slice.

Coverage: 4 new tests in
`tests/NetXlsx.Tests/NamedStyleTests.cs` — names + properties
survive file round-trip, case-insensitive lookup preserved,
workbook without named styles reads back empty (Excel's built-in
"Normal" filtered out), RegisterStyle replacement updates the
OOXML entry in place.

Test totals: 605 unit (was 601 at v1.2.0); other suites unchanged.

## [1.2.0] — 2026-05-22

v1.2 closes five of the six deferments from v1.1 and lands one
substantive refactor flagged by the v1.1 external review.

- **Slice 1 (refactor):** `ISheet.cs` partial-class split — 901 LOC
  monolithic file becomes 6 files split by abstraction level
  (IWorkbook / ISheet / ICell / IRow / IRange / IColumn). Mirrors
  the established `XssfCell.*.cs` / `XssfSheet.*.cs` impl-side
  pattern. Zero behavioral change.
- **Slice 2 (I-63):** `ISheet.RemoveTable(ITable)` — drops the
  OOXML `<tablePart>`, package relationship, and cached entry from
  `XSSFSheet.tables`. NPOI 2.7.3's `XSSFSheet` does not expose
  `RemoveTable` publicly; reaches across via reflection scoped to
  one method + one field, centralized in `Internal/NpoiInternals.cs`.
- **Slice 3 (I-64):** Per-column totals row on `ITable`. New
  `TotalsRowFunction` enum, `AddTotalsRow` / `RemoveTotalsRow` /
  `SetColumnTotal` / `SetColumnTotalLabel`. Uses 100-series SUBTOTAL
  formula codes for AutoFilter-aware totals; writes both the OOXML
  metadata AND the actual cell formula for cross-viewer robustness.
- **Slice 4 (I-65):** `IWorkbook.ProtectWithPassword(password, options?)`
  — writes the 16-bit XOR verifier into `CT_WorkbookProtection.workbookPassword`.
  Separately-named method rather than overload of `Protect` to
  sidestep call-site ambiguity + Roslyn `RS0027`.
- **Slice 5 (I-66):** Per-column AutoFilter criteria. New
  `FilterCriteria` class with 11 factories + `And` / `Or` /
  `Between` combinators, `ISheet.SetAutoFilterColumn` /
  `ClearAutoFilterColumn`. Covers Excel's custom-filter variant
  (equality, ordering, contains / startsWith / endsWith with
  wildcard escaping). Explicit-value list + Top-N variants
  deferred to v1.3 — NPOI 2.7.3 doesn't surface those properties
  on `CT_FilterColumn`.
- **Slice 6 (deferred to v1.3):** OOXML named-style table
  integration. Scope assessment during v1.2 surfaced that the
  full write + read integration is substantial enough to warrant
  its own focused slice; deferred rather than land a partial
  measure. v1.1 in-process behavior (decision I-57) remains
  correct and documented.

Decisions added: I-63 through I-66.

Test totals: **601 unit + 35 golden-file + 1 public-API snapshot +
18 fuzz = 655/TFM × 2 TFMs = 1,310 total** runs per CI build
(was 601/TFM at v1.1). PublicAPI.Shipped.txt = 569 entries
(was 535 at v1.1; +34 v1.2 surface).

### v1.2 slice 5 — per-column AutoFilter criteria (I-66)

Closes the v1.1 slice-7 deferment on AutoFilter. Adds the custom-
filter operator surface: 1-2 conditions per column joined by AND
or OR, covering equality, ordering, and string-pattern (contains /
startsWith / endsWith via Excel's wildcards).

- New `FilterCriteria` sealed class with 11 factories and two
  combinators (`And` / `Or`).
- New `ISheet.SetAutoFilterColumn(columnOffset, criteria)` —
  0-based offset within the AutoFilter range, matches OOXML's
  `colId`.
- New `ISheet.ClearAutoFilterColumn(columnOffset)`.
- Wildcard escaping: literal `*`, `?`, `~` in `Contains` /
  `StartsWith` / `EndsWith` arguments are escaped with `~` prefix
  before being wrapped in Excel's wildcard pattern.
- **Deferred to v1.3+**: explicit-value list filter (`filters`
  element, `In(...)` factory), Top-N filter (`top10` element),
  date-group filter, dynamic filter, color filter. NPOI 2.7.3's
  `CT_FilterColumn` doesn't surface those properties; would
  require XML-node-level workarounds.

Coverage: 16 new tests in
`tests/NetXlsx.Tests/AutoFilterTests.cs` — validation
(requires-AutoFilter, range bounds, null), single condition,
Between two-condition AND, OR combinator, three string-pattern
Theory cases, replace-on-same-column, ClearAutoFilterColumn
selectivity, And-chain limit at 2, wildcard escaping, file
round-trip.

Public-API snapshot + disposed-workbook matrix updated.

### v1.2 slice 4 — workbook password protection (I-65)

Closes the v1.1 slice-5 deferment. NPOI 2.7.3 did not expose a
workbook-level password helper; v1.2 writes the 16-bit XOR
verifier directly into `CT_WorkbookProtection.workbookPassword`.

- New `IWorkbook.ProtectWithPassword(password, options = null)`.
  Computes the verifier via NPOI's public
  `CryptoFunctions.CreateXorVerifier1`, encodes as 2-byte
  big-endian, writes to the CT_WorkbookProtection element.
- Same UX-guard-not-security caveat as I-54: the XOR-verifier
  algorithm is widely known to be brute-forceable.
- **Method-naming decision** (separate name vs overload of
  `Protect`): a same-named overload would create the call-site
  ambiguity `wb.Protect()` can't resolve, and an overload
  without defaults would violate Roslyn analyzer `RS0027`.
  Separate naming is also more self-documenting at the call
  site.

Coverage: 6 new tests in
`tests/NetXlsx.Tests/WorkbookProtectionTests.cs` — verifier-bytes
shape (`hunter2` → `C2 58`), default options, explicit options,
null-password rejection, Unprotect interaction with the password
byte array, and file round-trip.

Public-API snapshot + disposed-workbook matrix updated.

### v1.2 slice 3 — per-column totals row on `ITable` (I-64)

Closes the second v1.1 deferment. The v1.1 slice-2 surface (decision
I-51) made `ITable.HasTotalsRow` read-only — totals required per-column
function selection and didn't fit in slice scope. v1.2 ships the
full surface.

- New `TotalsRowFunction` enum (None / Sum / Min / Max / Average /
  Count / CountNumbers / StdDev / Var / Custom).
- New `ITable.AddTotalsRow()` — extends the table range by one row,
  flips `HasTotalsRow` to true, trims any AutoFilter range to
  exclude the totals row (Excel's default behavior).
- New `ITable.RemoveTotalsRow()` — clears per-column functions /
  labels, blanks the cells in the table's column range, shrinks
  the table range back by one row.
- New `ITable.SetColumnTotal(name, function)` — writes both the
  OOXML metadata and the SUBTOTAL formula into the cell. Uses
  the **100-series SUBTOTAL codes** (101..110) which skip
  AutoFilter-hidden rows. Uses the structured-reference form
  `TableName[ColumnName]` in the formula so totals auto-update
  when rows are added.
- New `ITable.SetColumnTotal(name, customFormula)` for arbitrary
  formulas. Strips a leading `=` if present.
- New `ITable.SetColumnTotalLabel(name, label)` — typically used
  for the leading "Total" cell. Explicitly clears the column's
  function metadata so the label takes precedence in Excel.

Out-of-scope (deferred): column names containing structured-
reference special characters (`#`, `[`, `]`, `'`, `@`, whitespace)
would need quoting in the formula body. v1.2 covers the
unquoted-name common case; names that need quoting reach through
`Underlying`.

Coverage: 20 new unit tests in
`tests/NetXlsx.Tests/TableApiTests.cs`, including a Theory matrix
verifying all eight built-in SUBTOTAL codes (101 / 102 / 103 /
104 / 105 / 107 / 109 / 110), validation rejections, label
override behavior, and a file round-trip preserving totals across
`Workbook.Open`.

Public-API snapshot + disposed-workbook matrix updated.

### v1.2 slice 2 — `ISheet.RemoveTable` (I-63)

Closes the first non-refactor item from the v1.2 roadmap. v1.1
slice 2 deferred this because NPOI 2.7.3's `XSSFSheet` didn't
publish a `RemoveTable` method; v1.2 implements it via the
three-step CT-part + relationship + cache cleanup pattern
already validated in v1.1 for other NPOI-internals workarounds.

- New `ISheet.RemoveTable(ITable)`. Drops the OOXML
  `<tablePart>`, removes the package relationship, and clears
  the cached entry from `XSSFSheet`'s internal `tables`
  dictionary.
- New `src/NetXlsx/Internal/NpoiInternals.cs` — centralized
  reflection over NPOI's protected `POIXMLDocumentPart.RemoveRelation`
  and private `XSSFSheet.tables` field. `MethodInfo` /
  `FieldInfo` cached as `static readonly` so each lookup
  happens once at class-init. Throws a clear "NPOI 2.7.3
  internal change" diagnostic if a future version moves either
  member, surfacing the breakage instead of failing silently.
- Validation surface: rejects null, rejects table handles that
  don't belong to this sheet (no matching relationship id), and
  rejects already-removed handles for the same reason.
  A second `RemoveTable(t)` on a freshly-removed handle throws
  loudly — not silently idempotent.
- After removal the codename is freed; a subsequent `AddTable`
  with the same name succeeds. Verified by file round-trip.

Coverage: 6 new tests in `tests/NetXlsx.Tests/TableApiTests.cs`.
Disposed-workbook matrix updated; `RemoveTable` exercised via an
inline stub.

## [1.1.0] — 2026-05-22

v1.1 ships the 10-slice "common asks" feature push from the v1.1
roadmap, plus four post-features improvements driven by the v1.0
external review and the new fuzz harness:

- **Style-pool diagnostics** — operational visibility (review #3)
- **Extended benchmark breadth** — micro + macro + percentile
  reporting (review #2)
- **NPOI 3.x evaluation checkpoint** — cadence holds, next
  re-check 2026-08-16 (review #4)
- **Fuzz harness + Open-path hardening** — closes the v1.1
  roadmap's fuzz-harness item; first run found and fixed an
  `IndexOutOfRangeException` leak from NPOI's parser

The "Unreleased" section above this header captures the full
slice-by-slice narrative. PublicAPI.Unshipped.txt flipped to
PublicAPI.Shipped.txt at this tag — 155 new public-surface
entries are now part of the SemVer-protected contract.

Decisions added in v1.1: I-50 (rich text), I-51 (tables), I-52
(images), I-53 (sheet protection), I-54 (workbook protection),
I-55 (data validation), I-56 (AutoFilter), I-57 (named styles),
I-58 (custom type converters), I-59 (strict concurrency
detection), I-60 (fuzz harness + Open-path hardening), I-61
(style-pool diagnostics), I-62 (extended benchmark coverage).

Test totals: **547 unit + 35 golden-file + 1 public-API
snapshot + 18 fuzz = 601/TFM × 2 TFMs = 1,202 total** runs per
CI build (was 868/TFM at v1.0).

### v1.1 cookbook recipes (release-PR prep)

Seven new cookbook recipes covering the v1.1 feature surface,
plus matching golden-file smoke tests:

- `rich-text-cells` — multi-run RichText in a release-announcement
  sheet (decision I-50).
- `excel-tables` — structured table with `TableStyleMedium2` +
  standalone AutoFilter on a sibling sheet (decisions I-51, I-56).
- `embedded-images` — PNG via magic-byte auto-detect, JPEG via
  explicit `ImageFormat` (decision I-52).
- `protected-template` — fully-locked reference sheet + partially-
  locked inputs sheet + workbook structure lock (decisions I-53, I-54).
- `validated-input-form` — list / integer / date / text-length /
  custom-formula validations (decision I-55).
- `branded-styles` — three named styles (header/body/footer)
  applied across cells and ranges (decision I-57).
- `custom-list-converter` — `List<string>` round-trip through a
  user-defined `ICellConverter<List<string>>` (decision I-58).

Each recipe was smoke-tested end-to-end via
`dotnet run --project samples/NetXlsx.Cookbook --no-build -c Release -f net10.0 -- <name> /tmp/x.xlsx`
and the produced files re-opened cleanly via `Workbook.Open`.

Cookbook total now **20 recipes** (13 v1.0 + 7 v1.1) discoverable
via the BenchmarkSwitcher-style entry-point list in `Program.cs`.

### Extended benchmark coverage (post-v1.1-features; external-review item #2)

- **New `benchmarks/NetXlsx.Benchmarks/BenchmarksExtended.cs`** with
  three new benchmark classes covering the gaps the v1.0 review
  identified (decision I-62):
  - `MicroBenchmarks` — per-cell write timing for each scalar
    type, A1-parse, style-pool hit/miss
  - `MacroBenchmarks` — 500 sheets, 100K cells in one sheet,
    streaming 200K rows
  - `ReadMicroBenchmarks` — single-cell open + read
- New `CiConfigWithPercentiles` config emits P50/P95/P99 to the
  regression-gate JSON alongside mean/median — long-tail
  regressions now fire the gate even when the central tendency
  is stable.
- All extended benchmarks run under the existing 15%-threshold
  regression gate. Program.cs uses `BenchmarkSwitcher.FromAssembly`
  so the new classes are auto-discovered.
- Closes the v1.0 external-review recommendation #2 ("extend
  benchmark suite to micro + macro shapes with percentiles").

### Style-pool diagnostics (post-v1.1-features; external-review item #3)

- **`IWorkbook.GetStylePoolDiagnostics()`** + **`StylePoolDiagnostics`
  value struct** (decision I-61). Read-only counters over the workbook's
  `CellStyle` + font dedup pools — exposes `StyleHitCount` /
  `StyleMissCount` / `FontHitCount` / `FontMissCount` / `UniqueStyles` /
  `UniqueFonts` plus convenience `StyleDedupRatio` and `FontDedupRatio`
  properties.
- Snapshot by value — does not allocate, does not update after capture.
- Closes the v1.0 external-review recommendation #3 ("expose style-pool
  diagnostics for ops visibility").

### Fuzz harness for the open path (post-v1.1-features hardening)

- **New project `tests/NetXlsx.Fuzz/`** (xUnit, opt-in via `[Trait("Category", "Fuzz")]`) — 18 tests across 6 fuzzing strategies (garbage bytes, empty/junk zips, truncated content-types XML, billion-laughs XML expansion bomb, high-compression-ratio zip bomb, bit-flip mutations of a known-good baseline, 100-iteration bulk random sweep with 2-second per-call cancellation cap). Closes the v1.1 roadmap's "fuzz harness for the open path" item (decision I-60).
- **`Workbook.Open` hardening**: the initial harness run surfaced `IndexOutOfRangeException` leaking from NPOI's parsers on adversarial input. `IsKnownMalformedOpenException` now also translates `IndexOutOfRangeException`, `NullReferenceException`, `OverflowException`, and `ArgumentOutOfRangeException` to `MalformedFileException`. The user-visible contract on `Open` for bad input is now strictly the documented exception family.

### v1.1 — Strict concurrency detection (slice 10/10, v1.1 complete)

- **`WorkbookOptions.StrictConcurrencyDetection`**: opt-in real-lock
  mode (decision I-59). When `true`, every structural mutating path
  on `IWorkbook` takes a real per-workbook `Monitor` lock,
  eliminating the gap between the default opportunistic reentry
  counter and silent corruption from concurrent threads.
- Default `false` — single-threaded callers don't pay the lock cost.
- Trade-offs vs the default opportunistic counter (decision #43):
  - Strict mode: concurrent mutations serialize cleanly (no
    exception); same-thread reentrancy permitted (Monitor is
    reentrant); some throughput cost.
  - Default mode: concurrent mutation may throw
    `InvalidOperationException` opportunistically; same-thread
    reentrancy is rejected; zero lock overhead.
- The default-mode exception message now mentions
  `StrictConcurrencyDetection` so callers discover the option when
  they hit a contention throw.
- Strict mode does not make the workbook thread-safe for reads —
  concurrent reads of any kind remain undefined.
- Coverage: 6 new tests
  (`tests/NetXlsx.Tests/StrictConcurrencyTests.cs`) — default
  value, single-threaded mutation, same-thread reentrancy (strict
  mode permits, default mode rejects), 50 parallel AddSheet
  serialize cleanly under strict mode with all 50 sheets present,
  default-mode behavior under contention does not produce silent
  corruption.

**v1.1 status: feature-complete.** 10/10 roadmap items shipped
between 2026-05-22 commits e576b27 through this slice. Total +~111
unit tests added (from 411 at v1.0 → 522 unit / 541 total
including SourceGen + golden-file). Public-API surface grew by
~140 entries in `PublicAPI.Unshipped.txt` (will flip to
`Shipped.txt` at the v1.1 tag).

### v1.1 — Custom type converters for typed mapping (slice 9/10)

- **`ICellConverter<T>` interface** + **`ColumnAttribute.ConverterType`**:
  let a `[Worksheet]`-mapped property carry any type the user can write
  a converter for, escaping the generator's built-in scalar set
  (decision I-58).
- A configured `ConverterType` overrides the built-in
  `IsSupportedPropertyType` check — the property emits write/read calls
  through a cached `static readonly ICellConverter<T> s_conv_X` field
  on the generated extension class. One instance shared across all
  `AddRow` / `ReadRows` calls.
- Source-generator changes:
  - `WorksheetProperty` model gains `ConverterTypeFullName`.
  - `IsSupportedPropertyType` returns true when a converter is set,
    regardless of underlying type.
  - `FormatSetCall` / `FormatReadExpression` emit converter dispatch
    when a converter is configured.
  - Emit pipeline outputs cached converter fields at the top of the
    generated extension class.
- **Why not a workbook-level registry?** The source generator runs at
  compile time and can't see runtime registrations. Per-property
  attribute binds at code-emit time.
- Coverage: 6 new tests
  (`tests/NetXlsx.Tests/CustomConverterTests.cs`) via the in-process
  generator harness — emits cached field, routes Write/Read through
  converter, no NXLS0006 for converter-property, mixed
  built-in-plus-converter still works, type with no SpecialType
  (List&lt;string&gt;) is not skipped from emission.

### v1.1 — Named / reusable styles (slice 8/10)

- **`IWorkbook.RegisterStyle` / `GetRegisteredStyle` /
  `RegisteredStyleNames`** + **`ICell.ApplyNamedStyle`** /
  **`IRange.ApplyNamedStyle`**: register a CellStyle by name on
  the workbook, apply by name on cells/ranges (decision I-57).
- Names are case-insensitive. Re-registering an existing name
  replaces the definition.
- **In-process convenience only.** v1.1 named styles do not
  produce OOXML named-style table entries; `Workbook.Open` does
  not rehydrate the name map. Per-cell visual style is still
  preserved through the existing style-pool dedup (decision #4) —
  naming is purely a caller-side ergonomic. Real OOXML named-
  style table integration deferred to v1.2.
- Coverage: 13 new tests
  (`tests/NetXlsx.Tests/NamedStyleTests.cs`) — empty default,
  register + get, case-insensitive lookup, get returns null for
  unknown, replace existing, enumerate names, null/empty name
  rejection, null style rejection, cell + range apply, unknown
  name throws. Disposed-workbook matrix updated.

### v1.1 — AutoFilter (slice 7/10)

- **`ISheet.SetAutoFilter(a1Range)` / `ClearAutoFilter` /
  `HasAutoFilter` / `AutoFilterRange`**: standalone AutoFilter for
  ranges that aren't structured tables (decision I-56).
- `SetAutoFilter` replaces any existing AutoFilter on the sheet.
  `ClearAutoFilter` removes it (no-op if none is set).
- **Per-column filter criteria deferred**. Excel's filter criteria
  model (text contains, top-N, color, date range, custom
  expression) is rich enough that exposing it would be a
  significant v1.1 surface chunk. Callers reach through
  `ISheet.Underlying.GetCTWorksheet().autoFilter.filterColumn`.
- **NPOI surprise**: `CT_Worksheet.autoFilter` is a direct property
  in NPOI 2.7.3 with no `IsSetX` / `UnsetX` accessors. Clearing
  assigns `null` to remove the element. The auxiliary
  `_FilterDatabase` built-in name (created by NPOI's
  `SetAutoFilter`) is left in place when clearing — Excel
  tolerates a stale name pointing at an absent autoFilter.
- Coverage: 9 new tests
  (`tests/NetXlsx.Tests/AutoFilterTests.cs`) — initial state,
  set + record range, replace, clear + idempotent clear, null
  + invalid range rejection, single-cell range, file round-trip.
  Disposed-workbook matrix updated.

### v1.1 — Data validation (slice 6/10)

- **`ISheet.AddValidation(a1Range, DataValidation)`**: apply data
  validation rules to a cell range (decision I-55).
- New public sealed class `DataValidation` with 11 static
  factories covering Excel's most-used validation families:
  - **List** (dropdown): `List(params string[])`,
    `ListFromRange(formula)`
  - **Integer**: `IntegerBetween`, `IntegerEqual`,
    `IntegerGreaterThan`, `IntegerLessThan`
  - **Decimal**: `DecimalBetween`
  - **Date**: `DateBetween(DateOnly, DateOnly)`
  - **Text length**: `TextLengthAtMost`, `TextLengthAtLeast`
  - **Custom formula**: `Custom(formula)`
- `DateBetween` uses Excel's `DATE(yyyy,m,d)` formula form rather
  than a locale-specific date literal — validation survives a
  round-trip on machines with non-US date formats.
- **Deferred to v1.2/.Underlying**: time-of-day validation,
  "not between" / "not equal" operators, formula-driven
  decimal/integer constraints, error-style customization
  (Stop/Warning/Information), per-validation prompt + error
  messages.
- Coverage: 12 new tests
  (`tests/NetXlsx.Tests/DataValidationTests.cs`) — factory
  validation, single + multiple rule application, range +
  single-cell address, file round-trip, locale-stable date
  formula. Public-API snapshot + disposed-workbook matrix
  updated.

### v1.1 — Workbook protection (slice 5/10)

- **`IWorkbook.Protect` / `IWorkbook.Unprotect` / `IWorkbook.IsProtected`**:
  workbook-level structure / windows / revision locks
  (decision I-54).
- New public record `WorkbookProtection` (Structure, Windows,
  Revision) + `WorkbookProtection.Default` and
  `WorkbookProtection.LockStructure` static instances. `Protect()`
  with no args defaults to `LockStructure` (the common use case —
  "stop accidental sheet add/delete").
- Same UX-guard-not-security caveat as sheet protection (I-53).
- **Workbook password support deferred.** NPOI 2.7.3 does not
  expose workbook-level password APIs directly. v1.1 ships the
  unprotected-by-default flag flip; password protection requires
  reaching through `.Underlying` for now.
- Coverage: 10 new tests
  (`tests/NetXlsx.Tests/WorkbookProtectionTests.cs`) — record
  semantics, enable + disable, Protect with default options
  clears unspecified flags, idempotent Unprotect, file round-trip.
  Public-API snapshot + disposed-workbook matrix updated.

### v1.1 — Sheet protection (slice 4/10)

- **`ISheet.Protect` / `ISheet.Unprotect` / `ISheet.IsProtected`**:
  toggle sheet-level UI protection (decision I-53).
- New public record `SheetProtection` with 15 granular `Lock*`
  flags (FormatCells, FormatColumns, FormatRows, InsertColumns,
  InsertRows, InsertHyperlinks, DeleteColumns, DeleteRows,
  SelectLockedCells, SelectUnlockedCells, Sort, AutoFilter,
  PivotTables, Objects, Scenarios) mirroring NPOI 2.7.3's
  `XSSFSheet.Lock*(bool)` methods. Plus `SheetProtection.Default`
  (all permissive) and `SheetProtection.LockAll` (all restrictive)
  static instances.
- **Documented limitation**: sheet protection is a UX guard, not
  security. Excel's sheet-protection password is hashed with a
  weak, known-brute-forceable algorithm.
- **NPOI surprise**: `XSSFSheet.ProtectSheet(null)` in NPOI 2.7.3
  is the *unprotect* operation, not "protect without password".
  The no-password path manipulates `CT_SheetProtection` directly
  to mirror the non-null side effects. Captured in
  `implementation-notes.md`.
- Coverage: 12 new tests
  (`tests/NetXlsx.Tests/SheetProtectionTests.cs`) — record
  semantics (Default / LockAll / structural equality), enable +
  disable behavior with and without password, idempotent
  Unprotect, granular flag propagation, file round-trip for both
  passwordless and password-protected sheets. Public-API
  snapshot + disposed-workbook matrix updated.

### v1.1 — Image embedding / PNG + JPEG (slice 3/10)

- **`ISheet.AddPicture(a1Cell, data, format)` / `ISheet.AddPicture(a1Cell, data)`**:
  embed PNG and JPEG images anchored to a single cell at natural
  pixel size (decision I-52).
- New public types: `IPicture` (Sheet, Format, Underlying),
  `ImageFormat` enum (Png, Jpeg), `UnsupportedImageFormatException`
  (extends `WorkbookException`).
- Auto-detection from magic bytes — 2-arg overload reads the
  leading bytes (`89 50 4E 47 ...` PNG, `FF D8 FF ...` JPEG) and
  dispatches; unknown formats throw
  `UnsupportedImageFormatException`. 3-arg overload skips detection.
- Pictures rendered at natural pixel size via NPOI's
  `XSSFPicture.Resize()` — without it, NPOI anchors to a
  single-cell extent and visibly stretches the image.
- **Other formats (GIF/BMP/TIFF/EMF/WMF) and advanced anchoring
  (multi-cell, pixel offsets, alt-text, rotation) deferred** —
  reach through `IPicture.Underlying`. v1.1 covers the
  two-formats-and-natural-size 90% case.
- Coverage: 10 new tests (`tests/NetXlsx.Tests/PictureApiTests.cs`)
  — explicit format, PNG/JPEG auto-detection, null + invalid-cell
  rejection, unknown-format detection error, file round-trip via
  `Workbook.Open`, multi-picture coexistence. Public-API snapshot
  + disposed-workbook matrix updated.

### v1.1 — Excel Tables / ListObject (slice 2/10)

- **`ISheet.AddTable` / `ISheet.Tables` / `ISheet.TryGetTable`**:
  sheet-scoped structured tables with a header row, optional style,
  and OOXML-mandatory AutoFilter (decision I-51).
- New public interface `ITable` (Name, DisplayName, Address, Sheet,
  ColumnNames, HasTotalsRow, StyleName, Underlying) + `TableStyles`
  static class with curated style-name constants (Light1, Light9,
  Light15, Medium2, Medium9, Medium16, Dark1, Dark9).
- `AddTable` validates: range has ≥2 rows; every header cell is a
  non-empty string; column names unique within the table; table
  name follows Excel rules and is unique workbook-wide (shares the
  namespace with named ranges).
- **`RemoveTable` deferred to v1.2.** NPOI 2.7.3's `XSSFSheet` has
  no `RemoveTable` method; package-part manipulation needed and
  not worth the v1.1 surface bloat. Reach through `.Underlying`
  for removal.
- **Per-column totals deferred to v1.2.** `HasTotalsRow` is
  read-only in v1.1.
- Coverage: 14 new tests (`tests/NetXlsx.Tests/TableApiTests.cs`)
  — happy path, style application, validation matrix (null/empty
  name, invalid name characters, A1-collision, duplicate
  workbook-wide, named-range collision, single-row range, empty
  headers, non-string headers, duplicate headers), and file
  round-trip via `Workbook.Open`. Public-API snapshot +
  disposed-workbook matrix updated.

### v1.1 — rich text in cells (slice 1/10)

- **`ICell.SetRichText(RichText)` / `ICell.GetRichText()`**: multi-run
  formatted strings (decision I-50). Per-run typography is restricted
  to font axes (Bold, Italic, Underline, FontName, FontSize, Color) —
  Excel's OOXML run model has no per-run fills/borders/alignment, so
  exposing the full `CellStyle` on a run would silently drop those
  axes. Cell-level style continues to route through `ICell.Style`.
- New public value records: `RichText`, `RichTextRun`, `RichTextStyle`
  (font-only subset of `CellStyle`; same nullable-axis convention).
- Run fonts are pooled through the existing `CellStylePool` font cache
  (decision #4) — runs with identical font properties share one
  `IFont` across the workbook.
- **`IStreamingCell.SetRichText` is intentionally absent.** NPOI's
  SXSSF `SheetDataWriter` (NPOI 2.7.x) reconstructs a fresh
  `XSSFRichTextString` from `cell.StringCellValue` at flush time,
  dropping all in-memory formatting runs. Per decision #7
  (streaming type-honesty), the absence of the method mirrors the
  absence of the capability rather than silently degrading.
  Caught during the v1.1 slice via a streaming round-trip test that
  surfaced the loss; documented in `docs/implementation-notes.md`.
- Coverage: 14 new tests (`tests/NetXlsx.Tests/RichTextApiTests.cs`)
  — value-type semantics, validation, in-memory round-trip, file
  round-trip via `Workbook.Open`, `MaxCellTextLength` enforcement,
  zero-length run skip, and the streaming-cell type-honesty
  reflection assertion. Public-API snapshot + disposed-workbook
  matrix updated.

## [1.0.0] — 2026-05-20

### ⚠️ BREAKING CHANGES

- **`net9.0` target removed.** Per decision **I24** (TFM support
  window policy), .NET 9 STS reached end-of-support on 2026-05-12
  and is dropped at this tag. `TargetFrameworks` is now
  `net8.0;net10.0`. Consumers on net9.0 will install the net8.0
  build via the standard TFM fallback — no API change, but worth
  knowing. Migrate to **net8.0** (LTS, supported through Nov 2026)
  or **net10.0** (LTS, current) at your convenience; both are
  fully supported.

### v1.0.0 — first stable release

The full pre-1.0 development arc — design discipline, pre-impl
spikes, decision log, public-API gating, golden-file preservation
fixtures, headless-no-fonts CI gate, benchmark regression gate —
all carried forward into this release. The library has been in
continuous test under three TFMs and across ubuntu + windows
runners since the public push.

**Highlights** (consolidated from the slice-level entries below):

- Workbook lifecycle, sheets, rows, ranges, columns with 1-indexed
  `[r, c]` access; fluent setters; `.Underlying` escape hatch on
  every public type for the 20% NetXlsx doesn't yet wrap.
- **Style-pool deduplication** (`CellStylePool`) — equal styles
  share one NPOI `ICellStyle` index. Avoids Excel's 64K-style cap
  that bites every team writing many-colored reports through raw
  NPOI. Measured as a correctness fix in spike 1.
- **Source-generator typed mapping** for `[Worksheet]`-decorated
  records — `sheet.AddRows<T>()` and `sheet.ReadRows<T>()` are
  emitted at compile time with no runtime reflection. AOT-safe by
  construction.
- **Type-level streaming split** — `Workbook.CreateStreaming()`
  returns `IStreamingWorkbook`, a separate interface from
  `IWorkbook`. Random-access members are absent from the
  streaming surface because they'd lie once a row is flushed past
  the window. (Decision #7, the "type-honesty" design choice most
  peer libraries get wrong.)
- **OPC preservation guarantee** (decision #44 / §7.7) — unmodeled
  parts (pivot caches, conditional formatting, custom XML,
  threaded comments) round-trip byte-identical through Open →
  Modify → Save. Verified by `RoundTripPreservationTests.cs`
  covering all four part categories.
- **Build-time AOT/trim guard** — setting `PublishAot=true` or
  `PublishTrimmed=true` produces MSBuild errors `NXLS0100` /
  `NXLS0101` rather than letting consumers discover the
  NPOI-side incompatibility at runtime.
- **`MissingFontException`** (decision I3) — `IColumn.AutoSize()`
  fails loud on headless hosts without a font stack rather than
  producing silently-wrong widths. Failure path verified by a
  dedicated `headless-no-fonts` CI gate.

**Test totals at tag time:** 434 tests/TFM × 2 TFMs = **868 total
runs per CI build** (405 unit + 28 golden-file + 1 public-API
snapshot per TFM). Bench gate active with rolling CI baseline.

**Public surface frozen.** `PublicAPI.Unshipped.txt` (380 entries)
flipped into `PublicAPI.Shipped.txt`; any post-1.0 addition has to
go through the normal Unshipped-then-Shipped-at-tag flow.

**Documentation snapshot at tag time:**
- `docs/design.md` — 52 foundational + 24 implementation decisions
  with rationale.
- `docs/roadmap.md` — binary v1.0/v1.1/v2.0/v3.0/Never matrix;
  per-version DoD; release-PR checklist.
- `docs/implementation-notes.md` — patterns and lessons from the
  pre-1.0 implementation phase.
- `docs/scheduled-spikes.md` — quarterly re-check cadence for
  NPOI AOT/trim posture (Spike 4-Q) and NPOI OSMF posture
  (Spike 5-Q).
- `docs/long-term.md` — post-v1.0 R&D direction; v2 OOXML path
  ordered by honest EV per the 2026-05-20 external critique.
- `docs/npoi-3x-migration.md` — concrete playbook for adopting
  NPOI 3.x once trigger conditions fire.
- `docs/v2-ooxml-planning.md` — research notes for the
  from-scratch path (option 4 of 4 in long-term.md's ordering).

Slice-level history follows below.

### Revise v2 planning docs per external critique (no code change)
A second external agent pointed out that the v2 planning docs led with
the from-scratch OOXML implementation as the headline option because
the maintainer had stated intent to "go all-in" on it. That was
commitment laundering, not planning. The revised framing puts
honest EV first.

- **`docs/long-term.md` adjacent options reordered** by expected
  value, not authorial preference:
  1. Bind ClosedXML as the engine (MIT, .NET-native, mature, no
     OSMF risk; reintroduces single-vendor-upstream pattern with
     a different vendor; does NOT unlock AOT).
  2. Fork NPOI 2.7.3 under Apache-2.0 (~30-50k LOC realistic
     maintenance surface, not the 200k headline — most is HSSF +
     HWPF we don't use).
  3. Accept OSMF terms after an honest legal read on transitive
     obligations to wrapper-library consumers.
  4. Full from-scratch — last, not first. Strengthens only if AOT
     becomes binding and neither NPOI 3.x nor ClosedXML moves on
     AOT in a reasonable horizon.

- **`docs/long-term.md` R&D sequence rewritten** as parallel
  spikes. R&D-1 now runs from-scratch *and* bind-ClosedXML
  prototypes at identical scope (single-sheet text+number write,
  ~2k LOC + 2-3 weeks respectively). Step 4 is a real gate that
  compares the two, not a continuation step. R&D-2 and R&D-3
  are reached only if the gate selects from-scratch. The
  original "R&D-1 → R&D-2 → R&D-3 from-scratch only" was a
  sunk-cost ladder that compared from-scratch against itself.

- **`docs/v2-ooxml-planning.md` reframed** to acknowledge it
  covers option 4 of 4, not the headline alternative.
  Hardening-phase estimate revised from "3-6 months" to
  "6-18 months full-time" (probably underweighted by 2-3× per
  the critique; ClosedXML's 427 open issues + 10 years of weird-
  file triage are evidence of the long-tail cost). Added explicit
  solo-maintainer-permanence caveat and opportunity-cost
  framing.

- **AOT centrality reframed** as the *active* question the
  quarterly Spike 5-Q watches for, not a "one case where I'd
  change my mind" contingency. New trigger conditions named:
  two-or-more of {NPOI-3.x-doesn't-do-AOT,
  ClosedXML-doesn't-do-AOT, consumer-side-AOT-demand} firing in
  the same quarterly review re-baselines from-scratch from
  option 4 to option 1.

- **Bus-factor honesty** added to `long-term.md`: the solo-
  maintainer cliff a 50k-LOC engine introduces is a delta on top
  of the wrapper's existing bus-factor, not a new category of
  risk. From-scratch makes it worse but doesn't introduce it.

- **Anchoring honesty paragraph** added to `long-term.md`:
  explicit acknowledgement that the original framing was skewed
  by maintainer's stated intent rather than honest EV. The
  revision removes the anchor.

The critique's biggest contribution was the parallel-spike R&D-1
design — that produces an actual decision input where the original
sunk-cost ladder didn't.

### Review actionables: NPOI 3.x plan, OOXML v2 planning, XssfCell split, test-count sweep
Six small items from the latest review's recommendations. No code
behavior changes; XssfCell refactored to partials with no public
surface change.

- **New doc `docs/npoi-3x-migration.md`** — concrete contingency plan
  for adopting NPOI 3.x once trigger conditions fire (license
  posture, API compatibility, AOT/trim re-spike, benchmark + test
  suite gating). Complements `docs/long-term.md` (which is the
  v2.0 "leave NPOI" path). Two docs, two paths, neither
  pre-committed.
- **New doc `docs/v2-ooxml-planning.md`** — research notes for the
  v2.0 path: ECMA-376 spec sizes + download locations, competitor
  project facts (ClosedXML / NPOI / Open-XML-SDK / EPPlus /
  MiniExcel — verified via GitHub API), realistic time estimate
  with confidence intervals (12-18 months solo full-time, 24-36
  months solo with day job), and a study sequence for the
  pre-implementation reading phase.
- **`docs/long-term.md` extended** with a "Roadmap re-baselining"
  section: post-v1.0, the non-Yes/non-Never matrix rows
  (v1.1/v2.0/v3.0/Deferred†) get a structured semi-annual review
  (Promote / Demote / Hold). Deferred† rows that don't promote in
  4 years auto-demote to Never. Prevents drift between major
  releases.
- **Roadmap: `WorkbookOptions.StrictConcurrencyDetection` v1.1
  entry** — reviewer-recommended opt-in that takes a real lock for
  callers who'd trade some throughput for "you cannot silently
  corrupt a workbook even if you ignore the thread-safety doc."
  Default stays opportunistic (decision #43 reentry counter).
- **`Internal/XssfCell.cs` split into four partial classes**
  (flagged by three consecutive reviewers). 495-LOC mega-file
  becomes:
  - `XssfCell.cs` (90 LOC) — core: fields, ctor, identity
    getters, Kind, Clear, `.Underlying`, default-style helper.
  - `XssfCell.Values.cs` (268 LOC) — SetX/GetX for every scalar
    type + formula + error code mapping.
  - `XssfCell.Style.cs` (63 LOC) — Style merge + NumberFormat +
    GetStyle + Merge helper.
  - `XssfCell.Annotations.cs` (113 LOC) — Comment + Hyperlink +
    SniffHyperlinkScheme.
  Same `internal sealed partial class XssfCell`; zero public API
  change; 434/TFM × 3 TFMs = 1,302 runs all green post-split.
- **Test-count sweep**: stale "433 tests/TFM" updated to **434**
  (the preservation fixture added one golden test). README,
  continuation file, and the v1.0 release-PR checklist's
  pre-drop/post-drop math all updated. Historical CHANGELOG
  entries left as-is (they accurately reflect the count at the
  time they were written).

### v1.0 ship-blockers all landed: AutoSize CI gate, bench regression gate, full preservation fixture
The three named v1.0 DoD ship-blockers from the latest review all
landed. v1.0 is now technically ready to tag (per the
release-PR checklist in `docs/roadmap.md`).

**Ship-blocker 1/3 — Headless-Linux AutoSize CI job (commit `66e4f4d`):**
- New `ColumnApiTests.AutoSize_Must_Throw_MissingFontException_When_NoFonts_Available`
  with `Trait("Category", "HeadlessNoFonts")` strictly asserts
  `MissingFontException` is thrown (no accept-either carve-out).
- Existing regular-CI matrix excludes this trait via
  `--filter Category!=HeadlessNoFonts`.
- New `headless-no-fonts` CI job in `ci.yml` runs on ubuntu-latest
  with all font packages + `libgdiplus` + `fontconfig` aggressively
  purged, then runs only the strict trait test. Verifies the design
  decision I3 promise that AutoSize fails loud on font-less hosts
  rather than silently producing wrong widths.

**Ship-blocker 2/3 — Benchmark regression CI gate (commits `97b981f`, `6b8be75`):**
- New `benchmarks/NetXlsx.Benchmarks/Benchmarks.cs` with five
  CI-friendly `[Benchmark]`s exercising the design §5 perf claims:
  `ColdCreateAndSave`, `Write5kRows`, `StyledWrite_SmallPalette`,
  `StreamingWrite_50kRows`, `OpenAndReadColumnSum`. Sized at
  ~45 seconds total via `ShortRun` config.
- New `benchmarks/compare-bench.py` reads BDN's brief-JSON
  output, compares per-benchmark `Statistics.Mean` to a baseline,
  exits 1 if any regresses > threshold (default 15% — design DoD's
  10% + 5% CI-noise headroom).
- New `.github/workflows/bench.yml` triggers on PRs and pushes
  that touch src, benchmarks, or build config. Caches a
  CI-hardware baseline keyed by source-tree hash; main pushes
  auto-refresh the baseline even when the run flags a regression
  (regression info on main is signal, not blocker); PRs fail loud
  on > 15% regression.
- Two-baseline model documented in `benchmarks/README.md`:
  committed `benchmarks/baseline/` is dev-local sanity reference;
  the CI cache is the actual regression gate.

**Ship-blocker 3/3 — Full preservation fixture (commit `4dfb001`):**
- `tests/NetXlsx.GoldenFiles/RoundTripPreservationTests.cs`
  expanded from synthetic-customXML-only to all four part types
  named in decision #44 / design §7.7:
  - Category 1: custom XML at `/customXml/item1.xml` (raw OPC).
  - Category 2: conditional formatting via NPOI's high-level API
    (greaterThan-50 rule on B1:B5, italic+bold font formatting) —
    serializes into the worksheet XML the way Excel would.
  - Category 3: pivot cache definition at
    `/xl/pivotCache/pivotCacheDefinition1.xml` (raw OPC stub with
    correct namespace).
  - Category 4: threaded comments at
    `/xl/threadedComments/threadedComment1.xml` (raw OPC stub
    with Excel 365 namespace).
- Three test methods: `All_Four_Unmodeled_Part_Types_Survive_Open_Modify_Save`,
  `Noop_Open_Save_Does_Not_Mutate_Any_Of_The_Four_Part_Types`,
  and the original single-customXML smoke test kept for back-compat.
- Fixture built programmatically per decision I18 option b (script-
  generated; inline helper acts as the `.gen.cs` sibling).

### Dep wave: NPOI 2.7.3 forward-compat, AwesomeAssertions, action bumps
Adjacent dependency hygiene that landed during the ship-blocker push:

- `BenchmarkDotNet` 0.14.0 → 0.15.8 (merged via PR #9).
- `actions/cache` v4 → v5 (latest stable; resolves a Node.js 20
  deprecation notice).
- `Internal/CellStylePool.cs`: switched `XSSFColor` construction from
  the `byte[]`-only ctor to `CT_Color`-based construction (commit
  `42fbda3`). Forward-compat to NPOI 2.7.4+, which removed the
  `byte[]`-only ctor — keeps our code resilient if we ever do bump
  to a 2.7.x patch, without taking such a bump now. Still builds and
  works on the pinned 2.7.3.
- PR #10 (NPOI 2.7.3 → 2.7.6) closed. The 2.7.x patch line has
  introduced two breaking API changes (2.7.4 removed `XSSFColor(byte[])`,
  2.7.6 `[Obsolete]`d `XSSFColor(CT_Color)`) — not patch-release
  discipline. Dependabot's NPOI patch updates are now also ignored
  alongside majors+minors; we opt in to specific bumps manually if
  upstream stabilizes.

### Doc tightening from review pass (no code changes)
Four documentation/policy refinements from the latest external review.
No code behavior changes; CI matrix unchanged.

- **AOT/trim matrix wording.** `roadmap.md` matrix rows for Native
  AOT and Trim compatibility now read **`Deferred†`** instead of
  `No†`. The footnote was expanded: two named paths to promotion
  (NPOI 3.x removes its problematic deps, *or* the native OOXML
  engine in `docs/long-term.md` lands), and an explicit note that
  AOT/trim are deferred-not-refused. The build-time MSBuild guards
  (`NXLS0100`/`0101`) still fail loud, but the matrix language now
  matches the actual project posture.
- **`Workbook.SanitizeSheetName` XML doc.** Now explicitly warns
  that sanitization can produce collisions (e.g. `"Q1/2026"` and
  `"Q1?2026"` both sanitize to `"Q1_2026"`) and points callers at
  `SuggestSheetName` for the sanitize-then-unique-against-workbook
  case. Closes the documented foot-gun the review flagged.
- **`docs/long-term.md` framing elevated** from "deferred aspiration"
  to **v2.0 R&D track**. The single OOXML-from-scratch section
  becomes a six-step milestone sequence (R&D-1: native write spike;
  R&D-2: native read spike; R&D-3: full coverage matrix run against
  the existing test suite; then a decision point). Each step is
  scoped enough to become a milestone with a date and an owner —
  not a free-form wishlist. No commitment to start; the work is
  recorded as gated on v1.0 stability and consumer signal.
- **v1.0 release-PR checklist** added to `roadmap.md` under
  "Process rules". Seven discrete steps the v1.0 release PR must
  execute: PublicAPI Unshipped → Shipped flip, PublicApiSnapshot
  baseline reconciliation, net9.0 TFM drop per I24, CHANGELOG
  breaking-change banner with migration guidance, version tag,
  `NUGET_API_KEY` secret verification, README/continuation/CHANGELOG
  test-count sweep, plus a confirmation that all three v1.0
  ship-blockers (benchmark CI gate, headless-Linux AutoSize CI job,
  round-trip preservation fixture) are landed.

Reviewer's three ship-blockers (benchmark regression gate, AutoSize
CI job, preservation fixture) remain as named v1.0 work — not
addressed in this slice; they're the next push, not a quick win.

### Add net10.0 TFM + patch CVE-flagged transitive deps (I22, I24)
- Target framework list expands to `net8.0; net9.0; net10.0` per
  decision **I22** (new TFMs added in the next minor release after
  GA). global.json bumps to require .NET 10 SDK. CI matrix and the
  release workflow install all three runtimes. AotSpike retargets
  to net10.0 to keep the spike on the newest available framework.
- New design decision **I24** records the TFM support window policy:
  **latest LTS + previous LTS + current STS** while in support.
  net9.0 STS support ended 2026-05-12; per I24 it will be dropped
  at the v1.0 tag (kept through pre-1.0 for adoption window).
  CHANGELOG will carry the drop notice when v1.0 lands.
- **Security**: the .NET 10 SDK's NuGetAudit surfaced three
  CVE advisories on NPOI 2.7.3's transitive deps that the .NET 9
  SDK didn't flag:
  - **GHSA-rxmq-m78w-7wmc** (moderate): `SixLabors.ImageSharp`
    2.1.10 → 2.1.11.
  - **GHSA-37gx-xxp4-5rgx** + **GHSA-w3x6-4m5h-cxqf** (both high):
    `System.Security.Cryptography.Xml` 8.0.2 → 8.0.3.
  Because NPOI itself is pinned at 2.7.3 (I23), the fixes apply via
  central package management's transitive pinning — explicit
  `PackageVersion` entries in `Directory.Packages.props`. Servicing
  releases on the same version lines — API surface unchanged.

Test count: still 433 per TFM (405 unit + 27 golden + 1 public-API).
Across the three TFMs that's now **1,299 total test runs** per build.

### Dependency hygiene: MinVer 7, AwesomeAssertions, NPOI 2.7.3 pin (I23)
Sweep of the four dependabot PRs opened on the initial push.
Outcomes summarized at the top, details below.

**Merged clean:**
- `MinVer` 5.0.0 → 7.0.0. No breaking change to our
  `MinVerTagPrefix` / `MinVerDefaultPreReleaseIdentifiers` config.
- `Microsoft.NET.Test.Sdk` 17.11.1 → 18.5.1
- `xunit` 2.9.2 → 2.9.3
- `xunit.runner.visualstudio` 2.8.2 → 3.1.5
- `coverlet.collector` 6.0.2 → 10.0.1
- `Microsoft.CodeAnalysis.Analyzers` 3.11.0 → 5.3.0 (the
  design-time analyzer package; safe in isolation)
- `Microsoft.CodeAnalysis.CSharp` 4.11.0 → 4.14.0 (latest 4.x;
  cannot bump to 5.x — see "Held" below)

**License-driven substitution:**
- `FluentAssertions` 6.12.2 → **removed**; replaced with
  `AwesomeAssertions` 9.4.0. FluentAssertions 8.0 switched to the
  Xceed Community License (free non-commercial only; commercial
  requires a paid license). AwesomeAssertions is the community
  fork from FA 6.12.2 under Apache-2.0. Namespace changed
  (`FluentAssertions` → `AwesomeAssertions` in `using` directives);
  `BeLessOrEqualTo` / `BeGreaterOrEqualTo` renamed to
  `BeLessThanOrEqualTo` / `BeGreaterThanOrEqualTo` — the only
  source-level breaking changes in our test code. All 433 tests
  per TFM still pass.

**Held — new design decision I23:**
- `NPOI` 2.7.3 → 2.8.0 **rejected**. NPOI 2.8.0 added an Open
  Source Maintenance Fee (OSMF) EULA on binary releases:
  organizations or users with ≥ US $10K annual revenue who depend
  on the library (directly or transitively) are required to pay
  a monthly maintenance fee. NetXlsx is MIT-licensed; passing the
  OSMF obligation transitively to downstream consumers would
  erode the "MIT all the way down" promise.
- New decision **I23** in `docs/design.md`: pin NPOI at 2.7.3
  (last clean Apache-2.0 release) and re-evaluate quarterly via
  the new `Spike 5-Q — NPOI OSMF posture re-check` in
  `docs/scheduled-spikes.md` (aligned with the existing AOT/trim
  Spike 4-Q cadence). Long-term direction recorded in the new
  `docs/long-term.md`: implement OOXML directly inside NetXlsx
  and drop the NPOI dependency entirely.

**Held — compiler-version compatibility:**
- `Microsoft.CodeAnalysis.CSharp` 4.x → 5.x **rejected**. 5.x
  requires Roslyn 5 (the .NET 10 SDK compiler); our TFMs are
  net8.0/net9.0 whose SDKs ship Roslyn 4.x. Source generators
  referencing a newer compiler than the loading `csc` fail with
  `CS9057`. Bumped to the latest 4.x (4.14.0) instead.

**Dependabot ignore rules added** (`.github/dependabot.yml`):
- `NPOI`: ignore major and minor bumps (I23). Patch updates
  within 2.7.x would still be welcome if upstream publishes any.
- `FluentAssertions`: ignore major+minor; the package is removed
  from the project.
- `Microsoft.CodeAnalysis.CSharp`: ignore majors. Patch updates
  within 4.x are fine.

Test count unchanged: 433/TFM (405 unit + 27 golden-file + 1
public-API snapshot).

### Pre-publish polish: CI, dependabot, contributor docs
Repo-side polish for the first public push. No library behavior changes.

Added:
- `.github/workflows/ci.yml` — build + test on push/PR. Matrix over
  `ubuntu-latest` and `windows-latest`. Installs `libgdiplus` +
  `fonts-dejavu-core` on Linux runners so `IColumn.AutoSize()` has a
  font stack and the AutoSize test takes the success branch (the
  test accepts either success or `MissingFontException` per decision
  I3, but green Linux CI now exercises the success path).
- `.github/workflows/release.yml` — on tag push matching `v*`,
  packs `NetXlsx` + `NetXlsx.SourceGen`, pushes to nuget.org
  (uses `NUGET_API_KEY` secret; skips push if absent), uploads
  the .nupkg artifacts to the workflow run, and creates a GitHub
  Release with generated notes. MinVer (already wired in
  `Directory.Build.props`) resolves the version from the tag.
- `.github/dependabot.yml` — monthly NuGet updates (grouped:
  test stack and analyzers) and weekly GitHub Actions updates.
- `CONTRIBUTING.md` — points contributors at the design doc, the
  public-API analyzer gate, the conventional-commits convention,
  the spike-before-design discipline, and the build entry points.
- `SECURITY.md` — vulnerability disclosure via GitHub private
  security advisories with a 90-day default coordinated-disclosure
  window. Calls out the SNK-in-repo as documented behavior (not
  a vulnerability) and routes NPOI-side findings upstream.

Cleaned:
- `nuget.config`: removed the placeholder feed-mapping comments
  that referenced a never-defined private feed.
- `CODEOWNERS`: collapsed the duplicate `@jkindrix @jkindrix`
  entries (artifact of an earlier dual-reviewer placeholder
  pattern) into a single default-owner line.

Removed:
- `.teamcity/settings.kts` — the project's CI lives in
  `.github/workflows/` for public OSS hosted on GitHub. Design
  §S17 updated to record GitHub Actions as the CI platform.

Repo presents on day one with a working CI pipeline, a release
path, dependency hygiene, and the standard OSS docs.

### v1.0-B — `WorkbookOptions` read-side safety + DisplayCulture-aware date rendering
Second half of the v1.0 `WorkbookOptions` slice. Closes the
read-side wiring. With this slice the v1.0 `WorkbookOptions`
contract is fully realized except for `DateSystem`, which v1
honors only informationally (NPOI hardcodes `date1904 = false`
on write — documented as a known constraint).

Behavior wired:
- `WorkbookOptions.ReadMaxSheets` — `Workbook.Open` rejects files
  whose `NumberOfSheets` exceeds the cap with
  `ResourceLimitExceededException("sheet count", limit, actual)`.
  Default is 1000 per design §6.1; well above any realistic file.
- `WorkbookOptions.ReadMaxUncompressedBytes` — best-effort
  post-open zip-bomb defense. Sums each OPC part's
  `GetInputStream().Length`; if the total exceeds the limit,
  throws `ResourceLimitExceededException("uncompressed package
  size in bytes", limit, total)`. Default 256 MiB.
  - NPOI's `PackagePropertiesPart` (core/extended/custom props)
    throws `"Operation not authorized"` on `GetInputStream()` —
    those parts are bounded-small and skipped in the sum.
  - True pre-buffer defense (inspect zip central directory
    before NPOI materializes) deferred past v1.0; documented
    in implementation-notes.
- `WorkbookOptions.DisplayCulture` — `XssfCell.GetString` on
  date-formatted numeric cells now routes through NPOI's
  `DataFormatter(culture)`, so date cells render per the
  configured culture (matches design §7.10). Bare numeric cells
  remain invariant G17 (§7.10 reserves culture-aware number
  rendering for v1.1+). Booleans never localize.

Known constraint (documented in implementation-notes):
- `WorkbookOptions.DateSystem` is informational only in v1.
  NPOI 2.7.x hardcodes `workbookPr.date1904 = false` in the
  `XSSFWorkbook` constructor; writing a 1904-epoch workbook
  isn't possible without reaching through the escape hatch.
  Read-side date interpretation is already correct because
  NPOI respects the file's own `IsDate1904()` flag.

Tests (+8): `WorkbookOptionsReadPathTests` covers
`ReadMaxSheets` (within-limit pass, over-limit
`ResourceLimitExceededException`, default-1000 doesn't reject
typical files), `ReadMaxUncompressedBytes` (within-limit pass,
1 KiB cap on a real .xlsx reliably trips the check), and
`DisplayCulture`-aware `GetString` (date cell renders
non-empty under both invariant and de-DE; bare number cell
stays invariant G17; bool stays invariant TRUE/FALSE).

### v1.0-A — `WorkbookOptions` entry-point wiring + write-side limit enforcement
First half of the v1.0 `WorkbookOptions` slice. The type shipped in
v0.9 but the random-access entry points ignored it; this slice wires
the write-side and default-font knobs through. Read-side safety
(`ReadMaxSheets`, `ReadMaxUncompressedBytes`), `DisplayCulture`-aware
`GetString`, and `DateSystem` land in v1.0-B.

Public surface (changes; no net additions to type count):
- `Workbook.Create(WorkbookOptions? options = null)` — new optional
  parameter. Existing `Workbook.Create()` calls continue to work.
- `Workbook.Open(string path, WorkbookOptions? options = null)` —
  same.
- `Workbook.Open(Stream stream, bool leaveOpen = true,
  WorkbookOptions? options = null)`.
- `Workbook.OpenAsync(string path, WorkbookOptions? options = null,
  CancellationToken ct = default)` — parameter inserted before
  `ct` per design §6.1.
- `Workbook.OpenAsync(Stream stream, bool leaveOpen = true,
  WorkbookOptions? options = null, CancellationToken ct = default)`.

Behavior wired:
- `WorkbookOptions.MaxCellTextLength` — `XssfCell.SetString` now
  reads the configured cap instead of a hardcoded `32_767`. Default
  matches the Excel hard cap, so callers see no behavior change
  unless they configure a smaller limit.
- `WorkbookOptions.MaxRowsPerSheet` — `XssfSheet.AppendRow`,
  `Row(int)`, and the `[r, c]` indexer now cap at
  `min(Options.MaxRowsPerSheet, CellAddress.MaxRow)`. Configuring
  a smaller value produces earlier failure with a message that
  reflects the configured cap.
- `WorkbookOptions.MaxColsPerSheet` — `XssfSheet.Column(int)`,
  `[r, c]` indexer, and `XssfRow.Cell(int)` cap the same way.
- `WorkbookOptions.DefaultFontName` / `DefaultFontSize` — applied
  in the `XssfWorkbook` ctor to the workbook's default font
  (NPOI font index 0). Defaults to Calibri 11 (matches Excel).
  Note: on the Open path, the file's authored default font is
  overwritten by these defaults unless the caller passes matching
  options — caveat documented in the new test.

Internal:
- `XssfWorkbook` gains a `WorkbookOptions Options` field, exposed
  to `internal` consumers (XssfCell, XssfSheet, XssfRow,
  XssfColumn, XssfRange). Construction takes the options via a
  new ctor overload; the no-arg overload defaults to
  `new WorkbookOptions()` and preserves the existing entry point.

Tests (+11): `WorkbookOptionsWritePathTests` covers null-options
equivalence, default-cap unchanged behavior, file round-trip with
options on both ends, `MaxCellTextLength` at both default (32,767)
and configured (50), `MaxRowsPerSheet` cap on `AppendRow`/`Row`/`[r,c]`
with a configured value of 5, `MaxColsPerSheet` cap on
`[r,c]`/`Row.Cell`/`Column` with a value of 3, default font Calibri
11, configured font Arial 14, and the open-default-font round-trip
caveat.

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
  GitHub Actions CI + release workflows, source project skeletons,
  test/benchmark/sample/golden-file project skeletons, public-API
  snapshot files).
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
