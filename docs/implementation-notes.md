# NetXlsx — Implementation Notes

A working file capturing patterns and lessons from the implementation
phase, with the project context they emerged in. This is **not** a
methodology — methodologies are written after the practices they describe
have been validated by completion. This file is the raw material from
which an implementation-phase methodology can be distilled after v1.0
ships.

Entries are dated and project-specific. Generalize at your own risk.

---

## 2026-05-15 — From v0.1.0 design lock to v0.2.0 first round-trip

### Vertical slice scoping

The reviewer pushed for a thin v0.2.0 vertical slice (Create → AddSheet →
SetString on `A1` → Save → Open → GetString) rather than building the full
§6 surface in one step. We did it. Outcome:

- The slice surfaced two genuine design questions the spec had not pinned
  down: how `XSSFWorkbook.Write` handles stream lifetime (NPOI closes by
  default; we override) and the exact form of `[MaybeNullWhen]` in the
  public-API analyzer file. Both were absorbed in 5-line fixes.
- 63 tests passing, full round-trip working, in one focused commit.
- Compared to the alternative ("ship the full §6"), the slice gave the
  source generator's `[Obsolete]` decoration a real reason to exist — the
  generator's stub methods still throw because `ISheet.AppendRow` etc.
  haven't landed yet, which they wouldn't have in either approach.

**Pattern:** when the design surface is large, scope the first
implementation milestone to an end-to-end thread, not a horizontal slice.
The thread surfaces specification gaps that horizontal work cannot.

### Stub policy: throw vs `[Obsolete(error: true)]` vs absent

Three options for an API surface that exists in spec but isn't yet wired:

| Option | Failure surface | When to use |
|---|---|---|
| Absent (don't emit) | Compile error: method doesn't exist | When the absence is reversible and no consumer should depend on it yet |
| `throw new NotImplementedException` | Runtime crash | Almost never — silently compiles, fails late, very poor signal |
| `[Obsolete(error: true)]` on the method | Compile error CS0619 with explanatory message | When the method *signature* is committed and consumers should write code against it that breaks at the right time |

We landed on `[Obsolete(error: true)]` for the generator's emitted
extensions. Reviewer specifically called out the throwing-stub anti-pattern
in commit `6af2ea8`. Lesson: pick the loudest failure your consumer-facing
contract can support.

### Design-doc revisions during implementation

When implementation surfaced something the design hadn't anticipated, we
took one of three paths:

1. **Edit the design row and continue.** Used when the surfaced item was
   small enough that the rationale fit in a sentence (e.g., the `Underlying`
   property type for `IWorkbook` resolving to `XSSFWorkbook` concretely).
2. **Treat as a structural revision and pause.** Used when a spike result
   invalidated an architectural assumption (spike 4: AOT/trim
   incompatibility re-classified two roadmap rows from `TBD*` to `No`).
3. **Defer with explicit milestone marker.** Used when the gap was real
   but the slice's scope didn't include it (concurrent-mutation detection
   per decision #43 — documented as known limitation in v0.2.0 CHANGELOG).

The discipline: never the fourth option, which is silent default. Every
implementation-time decision either edits the design or is recorded in
CHANGELOG as a deferred item.

### The PublicAPI analyzer earns its keep

`Microsoft.CodeAnalysis.PublicApiAnalyzers` blocked compile every time a
public symbol was added without a matching entry in
`PublicAPI.Unshipped.txt`. This forced the author to *think* about the
public surface in the same commit that introduced it — couldn't drift.
The wider lesson: a quality gate that the toolchain enforces beats a
quality gate that the reviewer enforces.

### Review-during-implementation cadence

After each milestone commit, we ran the work through an external review.
Patterns observed:

- The first review of any new code pass is dense — substantive items.
- The second review on the same code is usually cosmetic.
- Reviewers occasionally flag false alarms based on file listings (bin/,
  TestResults/) that are actually gitignored. Verify with `git ls-files`
  before acting. (Now codified in design-phase-methodology.md §5.)

### Honest constraint admission > optimism

The AOT/trim spike (spike 4) was the most valuable single thing we did
this session. The original design row was "TBD pending spike, likely
trim-with-warnings, AOT-incompatible." Spike 4 measured: both AOT and
trim fail at runtime, full stop. The roadmap matrix now says `No†` with
a real footnote. Consumers who would have hit a runtime crash six months
into adoption now get a build-time error courtesy of
`buildTransitive/NetXlsx.targets`.

### Analyzer noise vs analyzer signal

`TreatWarningsAsErrors=true` plus `latest-recommended` analysis level
produced a dozen analyzer errors during the v0.2.0 work — most of which
were genuinely useful (`ArgumentNullException.ThrowIfNull`,
`ObjectDisposedException.ThrowIf`, `string.Contains(char)` over
`IndexOf`). Two were noise for our specific design (`CA1720` on enum
value `String`; `RS0026` on intentional optional-param overloads). Both
were suppressed in `.editorconfig` with explicit comments.

**Pattern:** keep the gate; suppress at the rule level (not the file
level) when the rule is wrong for your design and the suppression has a
written justification.

---

## 2026-05-15 — Cookbook recipes 1-2 (post-v0.2.0)

### Cookbook recipe as load-bearing executable spec

The reviewer's argument played out exactly as predicted: writing
`HelloWorkbook` against today's API was clean; writing `TabularExport`
against it was *clunky in a specific, informative way*. The clunkiness
is the spec gap: `sheet[$"A{r}"].SetString(...)` per cell, row index
arithmetic in a string interpolation, no way to express "I have a row;
fill its cells." The pattern: leave the clunkiness in, document why, let
it motivate the next slice's interface rather than guessing what to
build.

### Recipe class as shared library between sample and test

Each recipe is a `public static class Foo { public static Task Run(string path) }`.
The cookbook executable dispatches to `Run`; the golden-file test
project references the cookbook project and calls `Run` directly. No
code duplication, the sample IS the test fixture, and the test catches
any drift between "what the recipe demonstrates" and "what the recipe
actually does."

Subtle benefit: the test asserts content via *both* NetXlsx's own
read path and a direct NPOI `XSSFWorkbook(stream)` read. If we ever
ship a NetXlsx-write bug that produces files only NetXlsx can
re-open, the direct-NPOI assertion catches it.

### Integer-literal SetNumber ambiguity

`cell.SetNumber(42)` does not compile — ambiguous between
`SetNumber(double)` and `SetNumber(decimal)`. Both are equally-valid
implicit conversions from `int`. This is a real call-site footgun:
anyone writing the obvious thing hits a build error and has to pick
`42.0` or `(decimal)42` or `42d`.

Resolution (next slice): add `SetNumber(int)` and `SetNumber(long)`
overloads. They're a strict superset of the current API — additive,
backwards-compatible — and they remove the ambiguity for every
integer-literal call site. The decimal vs double choice remains
explicit for floating-point literals where the precision-loss policy
(decision #36 / §7.4) is load-bearing.

Recipe code currently uses `42.0` with an inline comment pointing to
this note.

### Cookbook recipes as canary for the analyzer set

CA1859 fired on `IReadOnlyList<T>` return for a synthetic-data
generator (suggesting `T[]` for perf). Inside a sample whose explicit
purpose is to show idiomatic NetXlsx usage, the noise:signal ratio
is borderline. We took the rule's suggestion (changed to `T[]`) because
the cookbook is also a test fixture — perf rules matter there. If
recipes grow to demonstrate idiomatic *consumer* patterns where
`IReadOnlyList` is preferred, we'd suppress CA1859 specifically in
`samples/`.

---

## 2026-05-16 — IRow slice (v0.3.x)

### Recipe-driven interface design worked exactly as predicted

The v0.2.0 `TabularExport` recipe used `sheet[$"A{r}"].SetString(...)`-style
per-cell addressing. That clunkiness pinned the `IRow` contract:
`AppendRow()` for "add a row at the next index without computing it,"
`Set(col, value)` keyed by column for "fill this row's cells in order
of their column," and chainable fluent return type so a single record's
data is one statement. None of that came from staring at the design
doc; it came from looking at the recipe and asking "what would let me
write this cleanly?"

Diff from v0.2.0 to v0.3.x in TabularExport.cs is the readable record:
six lines of `sheet[$"A{rowNumber}"].SetString(r.Region); ...` per
record collapsed into one `sheet.AppendRow().Set(1, r.Region)...`
chain.

### Source generators don't flow transitively via ProjectReference

The Cookbook project (which now has its own `[Worksheet]` type) needed
its own `ProjectReference` to the generator with `OutputItemType="Analyzer"`.
Without it: build error CS1061 "ISheet has no AddRow." With it: works.

Downstream NuGet consumers get the generator automatically because
NetXlsx.csproj packs it under `analyzers/dotnet/cs/` in the nupkg.
But in-repo references need explicit per-consumer wiring. This is the
exact "cross-assembly types are invisible" footgun from decision I5,
hit from the opposite direction.

### Visibility = Public on [Worksheet] for cross-project consumption

The default `WorksheetVisibility.Internal` makes the emitted
`{Type}_SheetExtensions` class `internal`. Fine when the consumer of
the extensions lives in the same assembly as the `[Worksheet]` type.
Not fine when a test project in a different assembly wants to call
`sheet.AddRow(record)` against a recipe-defined `[Worksheet]` type.

The `SalesRecord` recipe type uses `[Worksheet(Visibility = WorksheetVisibility.Public)]`
explicitly. The recipe code comment makes the rationale visible to
readers.

### Roslyn pipeline cache: SpecialType is safe, ITypeSymbol is not

The generator needed to dispatch on the property's underlying type
when emitting cast text. First attempt: switch on `FullTypeName` (the
fully-qualified display string). Worked but format-dependent; would
break silently if the SymbolDisplayFormat ever changed.

Second attempt: store `ITypeSymbol` directly. Compiles but breaks
incremental caching — `ITypeSymbol` is a Roslyn handle, not value-
equal across recompilations.

Final: store `SpecialType` (Roslyn's special-type enum). Stable,
value-equal, exhaustive enough for v0.3.x's supported scalar set.

### Type narrowing/widening in emitted call sites

`IRow.Set` is overloaded only on the six core scalar types. A property
of type `short` is implicitly convertible to `int`, `long`, AND `double` —
ambiguous. The generator handles this by emitting an explicit cast
based on the property's `SpecialType`:

| Property type     | Emitted call form                |
|-------------------|----------------------------------|
| `int` / `bool` / `string` / `long` / `double` / `decimal` | `row.Set(col, record.X)` |
| `short` / `byte` / `sbyte` / `ushort` | `row.Set(col, (int)record.X)` |
| `uint` / `ulong`  | `row.Set(col, (long)record.X)`   |
| `float`           | `row.Set(col, (double)record.X)` |

`ulong` casts to `long` lossily for values > `long.MaxValue` — wraps.
Documented as v0.3.x scope; if a real use case wants overflow detection,
the generator can emit `checked((long)record.X)` instead. None today.

### v0.3.x scope-tightening on supported property types

`DateTime`, `DateOnly`, `TimeOnly`, `TimeSpan`, `Guid`, and all
`Nullable<T>` value types now trip `NXLS0006`. They were previously
declared supported (truthfully, per the design's long-term intent) but
the generator couldn't emit a valid Set call for them in v0.3.x because
`ICell` has no `SetDate/SetTime/SetDuration` overloads yet.

This is honest scope-shrinking: the design's promise stands; the
generator's capability today is narrower. NXLS0006 catches the
gap at compile time with a clear message. The supported list expands
when the corresponding `ICell` methods land.

---

## 2026-05-16 — Date / time / duration slice

### NPOI dates are numeric-with-format, not a distinct kind

NPOI stores `DateTime` as an IEEE-754 double (Excel's serial date
fraction) plus a number-format style. There is no "date type" at the
storage layer — only "numeric with a date-shaped format string."
`DateUtil.IsCellDateFormatted` is the classifier.

A real consequence: `TimeOnly` cells (stored as fraction-of-day with
format `h:mm:ss`) also classify as `CellKind.Date`, not
`CellKind.Number`. The first version of the golden-file test
expected `Number`; corrected to `Date` after the build surfaced the
assumption. The design's `CellKind.Date` definition is precisely
"numeric value styled with a date number format" — time formats
satisfy that, and the API is consistent.

### Lazy per-workbook style cache for the four default formats

`XssfWorkbook` holds nullable backing fields `_dateStyle`,
`_dateTimeStyle`, `_timeStyle`, `_durationStyle` and exposes them via
expression-bodied properties that materialize on first use. Four
styles, four properties, four format strings — and they're shared
across every cell that needs that format.

This isn't the full style-dedup pool the design specifies for v1.0
(`§4` perf targets, decision #4); it's a *targeted* cache for the
four formats NetXlsx applies automatically. The full dedup pool
lands with the styling API in a future slice. The two pools won't
conflict because the targeted cache only allocates styles whose
format strings are NetXlsx-controlled — they're stable inputs to
any future hash-keyed dedup.

### "Apply default style only if cell has no explicit style"

Decision I-18 says the workbook's default date format applies only
when the cell carries no explicit style. The implementation reads
the cell's current `Index`: index `0` is NPOI's workbook-default,
anything else is explicit. The detection is simple and matches the
spec; the test confirms a manually-set style is preserved when
`SetDate` is called afterwards.

The contract has one subtle gap: if the user sets a date-shaped
custom style FIRST, then calls `SetDate`, our code preserves their
style — even if they'd expected our default. This matches the
written rule but is worth knowing. In practice, the user setting a
style first is already opting into manual control.

### `TimeOnly` range is `[0, 1)` — anything else returns null

`GetTime` on a numeric cell with value `1.5` or `-0.25` returns
`null` rather than wrapping or throwing. The TimeOnly type covers
`00:00:00` through `23:59:59.9999999` — that's `[0, 1)` in
fraction-of-day terms. Values outside that aren't valid times of
day; returning `null` matches the design's "returns null for
non-convertible cells" contract.

`GetDuration` has no such restriction — `TimeSpan` happily holds
multi-day spans.

### Generator gained two non-special types via fallback path

`DateOnly`, `TimeOnly`, `TimeSpan` are not Roslyn `SpecialType`
values — they're regular types like any other. The generator's
`IsSupportedPropertyType` and `FormatSetCall` got fallback branches
keyed on the fully-qualified type name (`global::System.DateOnly`,
etc.). `DateTime` stayed on the `SpecialType.System_DateTime`
branch because it IS special.

This adds a second classification path through the generator. Worth
watching as more non-special types arrive (Guid will follow the same
pattern when its setter lands).

---

## 2026-05-16 — Styling slice (v0.4)

### Merge-on-apply, not replace

`ICell.Style(CellStyle)` merges the supplied style over the cell's
current style — non-null properties in the overlay replace null /
non-null on the existing axis. Two consequences:

- A previous `SetDate` keeps its number format when the user later
  calls `cell.Style(new CellStyle { Bold = true })`. The merged
  style is `{ Bold = true, NumberFormat = "yyyy-mm-dd" }`.
- The merge is strictly `existing ← overlay-non-null`, so passing
  `CellStyle.Default` (all-null) is a no-op — it does **not** clear
  the cell's existing style axes. Explicit clearing would require a
  sentinel ("set this axis back to null") which v0.4 does not have.
  Tracked as a known gap; minor. The XML doc on `ICell.Style`
  documents the merge semantics so callers aren't surprised.

The merge implementation is a single record `with`-style projection
in `XssfCell.Merge`. Twelve lines.

### Pool keys on CellStyle value equality

`Dictionary<CellStyle, ICellStyle>` keyed on `CellStyle`'s
auto-generated record `Equals`/`GetHashCode`. Worked first try
because `Color` is an ARGB-equal record struct and `CellBorders`
is an ordinary record — both contribute clean equality to the
parent `CellStyle`.

Watching for one risk: when an existing-style cell goes through
`ReadCurrentStyle` → `Merge` → `GetOrCreate`, the read step
reconstructs a `CellStyle` value record from the live NPOI
`ICellStyle`. If `ReadFromNpoi` introduces *any* extra property
(e.g., rounds `FontSize`), structural equality fails and the pool
allocates a fresh style. The first commit catches the common cases
(date defaults dedup correctly across cells); a stress test for
"100 cells, identical merged style, single ICellStyle index"
passes. Will be hardened with more property coverage as new axes
land.

### Font sub-pool

NPOI's `IFont` is itself a workbook-level allocated object; styles
share fonts when their font-relevant properties match. We mirror
that with a `Dictionary<FontKey, IFont>` so cells styled with the
same font properties (independent of whatever else varies in their
`CellStyle`) share a single `IFont`. Without this, 100 differently-
shaded-but-same-font cells would allocate 100 NPOI fonts — wasteful
and risks the 64K-style cap from a different angle.

### S29 cache absorbed cleanly

The "lazy four-format date-time cache" added in the v0.3.x date-time
slice (S29) became four pool entries. The four properties on
`XssfWorkbook` (`DateStyle`, `DateTimeStyle`, `TimeStyle`,
`DurationStyle`) now just call `StylePool.GetOrCreate(new CellStyle
{ NumberFormat = "..." })`. Zero changes to the cell-side call sites.
This is the methodology working: a documented interim that composes
cleanly with the full thing when it lands.

### Excel "fill foreground" vs "fill background" gotcha

NPOI's `FillForegroundColor` is the color shown when `FillPattern =
SolidForeground`. `FillBackgroundColor` is the *secondary* color for
patterned fills (cross-hatch, etc.). For a plain solid fill, you set
ForegroundColor + SolidForeground pattern. The first attempt wired
`FillBackgroundColor` and got nothing on cell display — fixed
silently in the same commit. Worth noting because the natural-English
naming pushes you the wrong way.

### 2026-05-16 — sub-slice C: column letter helpers promoted to public

Design §3 fixes the `IColumn` interface (with `Letter` and
`Column(string)` factory) but never spec'd the parser/formatter that
implements them — those lived as private helpers in `CellAddress`.
Sub-slice C promoted three of them to public:

- `CellAddress.ParseColumn(string)` — throwing form.
- `CellAddress.TryParseColumn(string, out int)` — non-throwing form.
- `CellAddress.FormatColumn(int)` — index → letter only (the
  single-letter form of the existing `Format(row, column)`).

Two reasons to expose rather than keep internal: (a) `IColumn.Letter`
is implemented by `FormatColumn`, and round-tripping `Column("XFD")`
→ `Index` → letter is a behavior callers will reasonably want to test;
(b) the parser is the same logic that backs `ISheet.Column(string)`,
so callers building dynamic column references benefit from the
non-throwing form rather than wrapping `try/catch`.

This isn't a design change — design §3 #284–285 implied these had to
exist — but the symbol-level additions are documented here because
they were not enumerated in design.md's CellAddress section (it
declares the type but doesn't enumerate members).

### 2026-05-16 — `WorkbookOptions.DateSystem` is informational in v1

Design §6.1 lists `DateSystem { Excel1900, Excel1904 }` on
`WorkbookOptions`. v1 honors it only **informationally**: the value
is stored on the workbook but not enforced or written to the
package, because NPOI 2.7.x's `XSSFWorkbook` constructor hardcodes
`workbookPr.date1904 = false` ("don't EVER use the 1904 date
system" — see NPOI source, `XSSFWorkbook.cs:435`).

For read-side date interpretation, NPOI already respects the
file's own `date1904` flag via `XSSFWorkbook.IsDate1904()`, so
opening a 1904 workbook authored by Mac Excel pre-2016 reads its
dates correctly with no v1 wiring required.

What this means in practice:
- Writing a workbook intended to use the 1904 epoch is not
  supported in v1. Callers needing that today must reach through
  `IWorkbook.Underlying` and mutate `CT_WorkbookPr.date1904`
  directly (deep NPOI internal).
- The `DateSystem` option on `WorkbookOptions` is kept in the
  public surface so the contract doesn't need to break when NPOI
  3.x (or a workaround) makes write-side 1904 viable.

### 2026-05-16 — `ReadMaxUncompressedBytes` is post-open, not pre-open

Real zip-bomb defense would inspect the central directory before
NPOI buffers any payload. NPOI's `XSSFWorkbook(Stream)` constructor
materializes the whole package in memory before returning control,
so v1's check runs *after* the buffer exists. This catches
over-the-line files at the fail-loud boundary (instead of allowing
them through to a downstream consumer), but it does NOT prevent
memory consumption during the open itself.

Two NPOI-side annoyances we work around in `EnforceReadLimits`:

- `NPOI.OpenXml4Net.OPC.Internal.PackagePropertiesPart`
  (core/extended/custom properties) throws
  `InvalidOperationException("Operation not authorized")` when
  you call `GetInputStream()` on it. These parts are bounded-small
  (a few hundred bytes of XML); we skip them in the size sum
  rather than fail the open.
- `PackagePart.Size` returns -1 by default and only a few
  subclasses override; we can't rely on it. The sum-by-stream
  path is necessary even though it allocates.

Pre-buffer zip-bomb defense would need OPC-level inspection
before NPOI parses (e.g., open the OPC package manually, read the
zip central directory, sum declared uncompressed sizes, reject if
over). That's a meaningful chunk of code and is deferred past
v1.0 — the current post-open check covers the "I opened a 256 MB
expanded payload by accident" case without blocking v1.

### 2026-05-16 — Streaming write: `IStreamingCell`, not `ICell`

Design §6.3 originally specified `IStreamingRow.Cell(int) -> ICell`,
with the comment "styling/value via the standard fluent ICell." That
contract can't be honored as written:

- `ICell.Underlying` returns `NPOI.XSSF.UserModel.XSSFCell` (concrete).
- The streaming cell type is `NPOI.XSSF.Streaming.SXSSFCell`, which
  does **not** inherit from `XSSFCell` — they share only the NPOI
  `ICell` interface.
- So a streaming-cell wrapper implementing `ICell` would either have
  to throw on `Underlying.get` (runtime trap) or return the wrong
  concrete type (compile-time lie).

Resolution shipped in v0.9 (decision **I-49**): a sibling
`IStreamingCell` interface, narrower than `ICell` — value setters,
`Style`/`NumberFormat`, `Kind`, address fields — with no
`Underlying`. Consumers reach the raw `SXSSFCell` through
`IStreamingSheet.Underlying.GetRow(idx).GetCell(col)` if they need
it. That's an honest reflection of streaming's narrower contract
(decision #7) and the constraint NPOI's type hierarchy imposes.

Design §6.3 updated to reflect this. The "Style merge semantics" on
`SxssfCell` are slightly weaker than on `XssfCell` because streaming
doesn't keep a reverse lookup from NPOI's `ICellStyle.Index` to a
`CellStyle` record for axes other than `NumberFormat` — incremental
`Style()` calls on the same cell may not preserve prior axes from
earlier calls. Practically a non-issue for the streaming use case
(bulk write, one styling pass), but worth knowing.

### 2026-05-16 — Named ranges: NPOI's name-uniqueness constraint

Decision I9 leaves cross-scope coexistence semantics unspecified.
Excel itself permits a workbook-scope name `Sales` *and* a
sheet-scope name `Sales` to coexist (sheet-scope wins when the
formula lives on that sheet). NPOI 2.7.x's `XSSFName.ValidateName`
rejects this — "The workbook already contains this name" — even
when the new name carries a non-default `SheetIndex`.

v1 therefore enforces workbook-wide uniqueness regardless of scope,
case-insensitively. Our own duplicate check is unconditional now
(not scope-aware) and produces a clearer message than NPOI's; we
keep the NPOI exception only as a backstop via the same wrapping
`ArgumentException` translation we use for invalid name/formula
parse failures.

Revisit if/when NPOI relaxes this (likely 3.x). Until then, the
contract documented on `IWorkbook.AddNamedRange` matches what we
can actually enforce, not Excel's full semantics. Tracked in
`docs/scheduled-spikes.md`-style follow-up — but the workaround for
callers who need cross-scope is to qualify name text explicitly
(`Sales_workbook`, `Sales_Data`).

Also discovered while writing tests: NPOI rejects names that parse
as cell references. `R1`, `C1`, `RC1` (R1C1-style) all fail with
"cannot be $A$1-style cell reference". Test fixtures use `Range1` /
`Range2` / `Range3` instead.

### 2026-05-16 — AutoSize font-failure translation

Decision I3 says "throw MissingFontException with installation
guidance"; the wire-up question is what to catch. NPOI 2.7.x's
column-sizing path goes through SixLabors.Fonts on netcoreapp, and
under WSL test runs we've never actually hit a font-missing failure
(DejaVu is installed by default on most Ubuntu base images). To
avoid swallowing real bugs, the `IsFontFailure` helper in
`XssfColumn` walks the inner-exception chain and matches only:

- Any type from `SixLabors.Fonts.*`.
- `System.Drawing.SystemFontsException` (legacy GDI+ path).
- `TypeInitializationException` (e.g. libgdiplus missing at type
  init).
- `IOException` / `FileNotFoundException` whose message mentions
  "font" or whose `FileName` contains "libgdiplus".

The `ColumnApiTests.AutoSize_Either_Sizes_The_Column_Or_Throws_…`
test deliberately accepts both outcomes: success (with `WidthUnits
> 0`) or `MissingFontException` (with a non-empty message). This
is the right shape until we have a dedicated headless-no-fonts CI
job — at which point that job can assert the throw, and the dev-box
run can keep asserting the success path.

### 2026-05-20 — NPOI 2.7.3 pin (decision I23)

NPOI 2.8.0 introduced the **Open Source Maintenance Fee (OSMF)**
model. The Apache-2.0 source remains in principle, but the binary
distribution on NuGet now carries an EULA: organizations or users
with ≥ US $10,000 annual revenue who depend on libraries in the
project family are obligated to pay a monthly maintenance fee via
GitHub Sponsors.

For NetXlsx — MIT-licensed, public OSS — accepting that bump means:

- Downstream consumers above the revenue threshold inherit a
  transitive monthly-fee obligation just by installing NetXlsx.
  Even if NetXlsx itself stays MIT, the package they receive
  pulls in a fee-encumbered binary.
- Our README's promise of "MIT all the way down" gets a footnote
  that erodes consumer trust at exactly the point most reviewers
  pay attention to (license details on a first-impression read).
- The OSMF terms aren't OSI-approved and the long-term legal
  interpretation (especially the transitive-via-wrapper question)
  isn't yet settled in public docs.

**Decision** (I23, 2026-05-20): pin at NPOI 2.7.3 — the last
Apache-2.0-only release. Dependabot configured (`.github/dependabot.yml`)
to ignore minor and major NPOI updates so the PRs don't keep
cycling. Quarterly re-check folded into the existing NPOI AOT/trim
spike cadence (`docs/scheduled-spikes.md`).

**Trade-offs accepted with this pin:**
- No upstream NPOI security fixes after 2.7.3. Mitigated by:
  the surface we expose is the wrapper, not raw NPOI; specific
  CVEs against NPOI 2.7.x can be patched by bumping to the
  patched 2.7.x if Apache-2.0-only patch releases continue, or
  by a NetXlsx-side workaround.
- No upstream bug fixes for the NPOI quirks we documented
  (workbook-name-uniqueness constraint, `date1904 = false`
  hardcode, etc.). Those stay as-is until either the pin moves
  or the long-term own-OOXML work lands.

**Forward direction.** The long-term resolution is to implement
OOXML directly inside NetXlsx and drop the NPOI dependency
entirely. That's a multi-month effort and a separate-from-v1.0
project; tracked in `docs/long-term.md`. v1.0 ships on the 2.7.3
pin and the wrapper architecture; the OOXML-from-scratch work
becomes a v2.0 conversation after v1.0 is in consumer hands and
the use cases have stabilized.

### 2026-05-20 — net10.0 added; transitive security fixes (I24)

Added `net10.0` to the TFM list per decision **I22** (net10.0 GA reached
Nov 2025; adding now in the next minor release). New list:
`net8.0; net9.0; net10.0`. global.json bumped to require the .NET 10
SDK as minimum (rollForward=latestFeature). CI + release workflows
install all three runtimes.

The .NET 10 SDK's NuGetAudit caught three CVE advisories that .NET 9
SDK either didn't flag or wasn't running against (the advisories
likely post-date the last .NET 9 build). All three are transitive
through NPOI 2.7.3:

- **GHSA-rxmq-m78w-7wmc** (moderate): `SixLabors.ImageSharp` 2.1.10
  vulnerable. Fixed in 2.1.11.
- **GHSA-37gx-xxp4-5rgx** + **GHSA-w3x6-4m5h-cxqf** (both high):
  `System.Security.Cryptography.Xml` 8.0.0–8.0.2 vulnerable. Fixed
  in 8.0.3 (and 9.0.15, 10.0.6 on the other release lines).

Because we can't bump NPOI itself past 2.7.3 (decision I23, OSMF),
the fixes are applied via central package management's transitive
pinning — explicit `PackageVersion` entries for both packages in
`Directory.Packages.props` at the patched servicing releases.
`CentralPackageTransitivePinningEnabled=true` (already set) makes
NPOI's transitive resolution honor our pins instead of its declared
versions. API surface is identical on the servicing releases, so no
behavioral change.

This is exactly the kind of finding that shipping to public OSS
*should* produce: the broader audit tooling .NET 10 SDK runs surfaced
real CVE exposure that the internal-only build never caught. Worth
naming as an unexpected positive of going public earlier rather than
later.

Decision **I24** records the TFM support window policy: latest LTS +
previous LTS + current STS. v1.0 cut will drop net9.0 (EOS 2026-05-12)
per that policy. CHANGELOG carries the drop notice.

### 2026-05-20 — FluentAssertions → AwesomeAssertions

FluentAssertions 8.0 switched to the Xceed Community License:
free only for non-commercial use; commercial use requires a paid
Xceed license. v6 and earlier remain Apache-2.0.

Same logic as the NPOI pin applies: a commercial-encumbered test
dependency on an MIT library hurts the "free for commercial use"
promise at exactly the point reviewers verify it.

**Decision** (informal — test-stack choice, not API-affecting):
migrate to **AwesomeAssertions** 9.4.0, the community fork that
continued from FluentAssertions 6.12.2 under Apache-2.0. Namespace
changed (`FluentAssertions` → `AwesomeAssertions` in `using`
directives); two assertion methods renamed (`BeLessOrEqualTo` →
`BeLessThanOrEqualTo`, `BeGreaterOrEqualTo` → `BeGreaterThanOrEqualTo`).
All 433 tests/TFM passed after the swap.

Dependabot configured to ignore FluentAssertions major bumps so
the 6.x → 8.x PR doesn't keep cycling.

---

## 2026-05-22 — v1.1 slice 1: rich text in cells (I-50)

### NPOI SXSSF drops rich-text formatting at the serializer level

The streaming round-trip test for `SetRichText` is what forced the
final shape of the surface. Set up:

1. Add `SetRichText` to both `ICell` (random-access) and
   `IStreamingCell` (streaming).
2. Implement on both `XssfCell` and `SxssfCell` via NPOI's
   `XSSFRichTextString.ApplyFont(start, end, IFont)`.
3. Write a streaming round-trip test: create → set rich text →
   save → open → assert.

The streaming round-trip test failed with `GetRichText() == null`
on read-back. Inspecting NPOI's `SheetDataWriter.cs:307-336`
revealed the cause:

```csharp
case CellType.String:
{
    if (_sharedStringSource != null)
    {
        XSSFRichTextString rt = new XSSFRichTextString(cell.StringCellValue);
        //                                              ^^^^^^^^^^^^^^^^^^^^
        // Plain string only — no formatting runs from the original cell value!
        int sRef = _sharedStringSource.AddEntry(rt.GetCTRst());
        ...
```

SXSSF's serializer reconstructs the rich-text string from the
plain string at flush time, **dropping all formatting runs**. So
`SxssfCell.SetCellValue(IRichTextString)` stores the runs in
memory, but the on-disk OOXML has none.

This isn't NetXlsx's bug — it's a known SXSSF limitation in NPOI
2.7.x. We had three options:

1. **Throw `NotSupportedException`** on
   `IStreamingCell.SetRichText` — honest, but creates an awkward
   throw-on-call surface.
2. **Silently degrade** to plain string write — quiet data loss,
   violates the spirit of the type-honest streaming split
   (decision #7).
3. **Omit the method from `IStreamingCell`** — absence of the
   method mirrors absence of the capability.

Picked (3). Decision #7 was the load-bearing precedent — random-access
`Workbook` and streaming `IStreamingWorkbook` are separate interfaces
specifically so streaming doesn't expose capabilities it can't honor.
This is the same principle applied one level down.

The reflection assertion in `RichTextApiTests` (`IStreamingCell_Does_Not_Expose_SetRichText`)
documents the absence and pins the decision into the test surface
so a future addition (e.g., NPOI 3.x reopening SXSSF rich-text
support) is a deliberate change, not a silent re-introduction.

### Font-only run style: rejecting the obvious choice

The first surface shape proposed `RichTextRun(string Text, CellStyle? Style)`
— reuse the existing `CellStyle` type for runs. Tempting, but wrong:

- Excel's OOXML `<r><rPr>` element has no fills, borders, alignment,
  number format, or wrap-text axes. Per-run, only font properties
  apply.
- Passing `CellStyle { Background = Color.Red }` on a run would
  silently no-op the background. Callers wouldn't see the drop.
- The type signature lies — accepts what looks like a complete
  style; honors a subset.

A purpose-built `RichTextStyle` with only the font axes (Bold,
Italic, Underline, FontName, FontSize, Color) makes the surface
honest. Cell-level style continues through `ICell.Style(CellStyle)`
unchanged.

This is one of those "the easy option is the wrong option" moments
the design discipline keeps catching. A function-only type
duplicates six properties from `CellStyle`, but the duplication
is load-bearing: it documents what the run model actually
supports.

### Read-back: distinguishing "rich" from "plain"

NPOI stores every string cell as `XSSFRichTextString` internally,
including cells set via plain `SetCellValue(string)`. `GetRichText()`
needs to distinguish:

- A cell set via `SetString("hello")` → return `null`.
- A cell set via `SetRichText(...)` → return the runs.
- A cell loaded from a file with explicit formatting → return runs.

The discriminator is `XSSFRichTextString.NumFormattingRuns`. Plain
strings have 0; explicit formatting has ≥ 1. Captured in the unit
tests so the distinction stays load-bearing.

### Font pool reuse

The existing `CellStylePool._fontPool` (keyed by `FontKey`)
already supports rich-text's needs perfectly — same fields
(name/size/bold/italic/underline/color). Added one internal method
`GetOrCreateRunFont(RichTextStyle)` that maps to `FontKey` and
delegates. No new pool, no new key type, no duplication.

Lesson: when a new surface needs structural infrastructure the
existing surface already has, look for an internal-method
extension before assuming a new pool/cache is needed.

---

## 2026-05-22 — v1.1 slice 2: Excel Tables / ListObject (I-51)

### NPOI 2.7.3's `XSSFTable.CreateColumn` throws on a fresh table

Wrote the obvious implementation first: `sh.CreateTable()` →
`t.Name = ...` → `t.CreateColumn(header)` per header → `t.SetCellReferences(area)`.
Tests failed with:

```
System.ArgumentOutOfRangeException : Index was out of range.
   at System.Collections.Generic.List`1.get_Item(Int32 index)
   at NPOI.XSSF.UserModel.XSSFTable.CreateColumn(String columnName, Int32 columnIndex)
```

Root cause: NPOI 2.7.3's `XSSFTable.CreateColumn` indexes into
`ctTable.tableColumns.tableColumn` without first checking that the
list itself is non-null. On a fresh `CreateTable()` return, that
inner `List<CT_TableColumn>` is null. The trunk source we read
earlier (which guards with `if (columns == null) columns = ctTable.AddNewTableColumns()`)
matches a later NPOI version, not 2.7.3.

This is a recurring kind of trap with NPOI: the trunk source on
disk diverges from the released NuGet, and the documentation .xml
is a sparse subset of the actual public surface. Specifically:

- `XSSFTable.CreateColumn` — public in 2.7.3 source, broken at call site.
- `XSSFTable.AddColumn` — exists in trunk source, **absent** from 2.7.3 binary.
- `XSSFTable.StyleName` setter — public in 2.7.3 (works), but not in .xml docs.
- `XSSFSheet.RemoveTable` — exists in trunk source, **absent** from 2.7.3.

**Workaround pattern adopted:** populate `CT_Table.tableColumns`
directly via the OOXML proxy types, bypassing `CreateColumn`:

```csharp
var ctTable = npoi.GetCTTable();
var ctColumns = new CT_TableColumns
{
    tableColumn = new List<CT_TableColumn>(headers.Count),
    count = (uint)headers.Count,
};
ctTable.tableColumns = ctColumns;
for (int i = 0; i < headers.Count; i++)
{
    var ctCol = ctColumns.InsertNewTableColumn(i);
    ctCol.id = (uint)(i + 1);
    ctCol.name = headers[i];
}
```

This bypasses the broken accessor entirely and produces output
that round-trips cleanly through `Workbook.Open`. Verified with a
standalone probe app against NPOI 2.7.3 before committing.

**General lesson:** when an NPOI call from the source-line reads
"obviously right" but throws at runtime, before debugging the call,
check the actual NuGet binary surface (`grep` the .xml docs, or
reflect via `typeof(T).GetMembers()`). Trunk source is aspirational;
the binary is what's deployed. The probe-app pattern (a throwaway
`/tmp/probe-app` `dotnet new console` referencing the same NPOI
version) is the fastest way to confirm.

### "RemoveTable" deferred — package-part manipulation isn't worth it

NPOI 2.7.3 has no `XSSFSheet.RemoveTable`. Removing a table from a
sheet requires:

1. Removing the `<tablePart>` from `CT_Worksheet.tableParts`.
2. Removing the relationship via `package.RemovePart()`.
3. Updating the sheet's part-loaded `tables` dictionary.

That's three coordinated mutations and zero observed user demand.
v1.0 shipped without table support at all, and the v1.1 user-ask is
"make tables", not "remove tables programmatically". Deferred to
v1.2 (or whenever NPOI 3.x lands). Documented in CHANGELOG +
design row I-51. Callers needing removal today reach through
`.Underlying` and do the three-step dance themselves.

### Name-space sharing between tables and named ranges

Excel treats table codenames and named-range names as a single
workbook-wide namespace. Two tables on different sheets cannot
share a codename, and a table cannot have the same name as a
named range. Reused `XssfWorkbook.GetAllNames()` enumeration to
catch the named-range collision; walked all sheets'
`GetTables()` to catch the cross-sheet table collision.

Friendlier error message than letting NPOI's serializer fail at
save time. Same pattern as the named-range uniqueness check from
v0.7.

---

## 2026-05-22 — v1.1 slice 3: image embedding (I-52)

### `XSSFPicture.Resize()` is the difference between visible and invisible

NPOI's `XSSFDrawing.CreatePicture(anchor, idx)` accepts an anchor
specifying the from-cell. If from-cell equals to-cell (single-cell
anchor), the picture is rendered at the cell's display size — for a
typical narrow Excel column, that's a barely-visible smear.

The fix is `XSSFPicture.Resize()` (no args), which sets the to-cell
based on the image's pixel dimensions. Excel then renders at the
image's natural size.

Without `Resize()`, AddPicture produces a technically-correct file
with a visibly-broken picture. Caught during initial test runs;
documented because it's not obvious from NPOI's surface (Resize()
sounds like an optional convenience, not a near-mandatory step).

### `picture.Resize()` requires a real image, not a hand-crafted byte buffer

First-attempt PNG fixture was hand-crafted (manually computed CRC
fields). NPOI's `Resize()` calls `SixLabors.ImageSharp` to decode
the image and read dimensions; ImageSharp validated the CRCs and
threw `InvalidImageContentException`. Lesson:

- For tests that exercise `AddPicture`, the byte buffer must be a
  *real* image that survives a proper PNG/JPEG decoder, not just
  one with the right magic bytes for our auto-detect.

Fix: embed the canonical 67-byte "tiny transparent PNG" via
`Convert.FromBase64String("iVBORw0KGgo...")`. Known to load in
every PNG decoder. Same approach for the JPEG fixture (smallest
valid 1×1 JPEG — about 125 bytes, hand-assembled from the
SOI/DQT/SOF0/DHT/SOS/EOI markers).

### Format detection: magic bytes, not file extension

The 2-arg `AddPicture(a1Cell, data)` could have looked up format
by file extension if we accepted a path. We took bytes-only, so
magic-byte detection is the right move:

- PNG: `89 50 4E 47 0D 0A 1A 0A` (8-byte signature)
- JPEG: `FF D8 FF` (3-byte SOI + first APP marker)

Anything else throws `UnsupportedImageFormatException` with a
message naming both supported formats. Avoids the silent
mis-classification that file-extension sniffing would invite.

---

## 2026-05-22 — v1.1 slice 4: sheet protection (I-53)

### `ProtectSheet(null)` means "unprotect", not "protect without password"

First-attempt implementation called `_underlying.ProtectSheet(password)`
uniformly, passing the caller's `password` through (including
`null`). Tests immediately failed for the no-password case:
`IsProtected` returned `false` after `Protect()`.

NPOI 2.7.3 source confirms (`XSSFSheet.cs:2579`):

```csharp
public void ProtectSheet(string password)
{
    EnsureWorksheetLoaded();
    if (password != null)
    {
        CT_SheetProtection sheetProtection = worksheet.AddNewSheetProtection();
        SetSheetPassword(password, null);
        sheetProtection.sheet = true;
        sheetProtection.scenarios = true;
        sheetProtection.objects = true;
    }
    else
    {
        worksheet.UnsetSheetProtection();   // <— null path removes protection
    }
}
```

The null branch is documented as "remove protection", not
"protect without a password." `Unprotect()` therefore happily
delegates to `ProtectSheet(null)`, but `Protect()` without a
password has to take a different path.

**Workaround:** mirror what the non-null branch does, minus the
password step — create `CT_SheetProtection` directly with
`sheet=true; scenarios=true; objects=true`. The granular `Lock*`
flags can then be set on top via the existing `LockX(bool)`
methods (they all assume the element exists).

Same family of NPOI gotcha as the `XSSFTable.CreateColumn`
surprise from slice 2: an API that *looks* like it should do
one thing actually does another, and the only tell is reading
NPOI source or running a probe app. The pattern is consistent
enough now that "probe-app first when NPOI seems wrong" is the
default reflex.

### "Sheet protection" is not security; say so in the XML doc

Excel's sheet-protection password uses a hash (CRC-style; in
2007+ it's iteratively SHA-stretched but the algorithm is still
documented and trivially brute-forced because the password
space is small). Tools like `msoffice-crypt`, `office2john`, and
even online "unprotect" services exist.

If we don't document this in the type's XML doc, callers who see
"password" assume real cryptographic protection. Added a
prominent caveat in `SheetProtection`'s class doc and in the
`ISheet.Protect` method doc. Same pattern as the
"escape-hatch direct mutation is not synchronized" caveat on
`.Underlying`: the surface allows the operation, the docs make
clear what guarantees apply.

---

## 2026-05-22 — fuzz harness drives Open-path hardening (I-60)

### The harness immediately found something

First run of the new fuzz harness produced a failing bit-flip test:
NPOI's parser threw `IndexOutOfRangeException` from somewhere inside
`XSSFWorkbook(stream)`, propagating past NetXlsx's
`IsKnownMalformedOpenException` filter. The runtime exception
escaped into user code as-is.

This is exactly the kind of finding the harness is meant to catch:
not a logic bug in NetXlsx, but a contract gap. Our `Open` documents
that bad input surfaces as `MalformedFileException`. A leaked
`IndexOutOfRangeException` violates that contract — the user can't
write a clean `try { Open(...) } catch (MalformedFileException) { ... }`
if NPOI bypasses it.

**Fix:** extend `IsKnownMalformedOpenException` to translate the
runtime-exception family commonly produced by parsers on truncated
or adversarial input:

```csharp
return ex is System.IO.InvalidDataException
    or System.IO.IOException
    or System.IO.EndOfStreamException
    or System.Xml.XmlException
    or System.IndexOutOfRangeException    // new
    or System.NullReferenceException      // new
    or System.OverflowException           // new
    or System.ArgumentOutOfRangeException // new
    or FormatException;
```

Critical exceptions (`OutOfMemoryException`, `StackOverflowException`,
`OperationCanceledException`) are still excluded — those indicate
runtime / programmer fault, not bad data, and propagate verbatim.

### Why translate rather than fix NPOI

In a perfect world, NPOI's parsers would also throw a typed exception
for "input doesn't match my schema." They don't, consistently. We can
either:

1. File NPOI bugs and wait (months/years; the project moves at its
   own pace).
2. Translate at our boundary.

Option 2 is the right v1.x posture — we own the surface, the
contract belongs to us, and the translation cost is one extra `or`
clause. The fuzz harness will keep finding new leaks; each one is a
small addition to the same filter.

### Harness design choices

**xUnit, not a separate fuzzer binary.** SharpFuzz / libFuzzer would
get more coverage per CPU-second, but they require an external
fuzzer process and a corpus directory checked in. The xUnit
harness sits in the same test infrastructure as everything else —
runs locally without setup, runs in CI with one extra line, and
the assertions are testable by code review.

**Opt-in via trait, not separate solution.** `[Trait("Category", "Fuzz")]`
lets CI gate the slow bulk-sweep test (100 iterations × up to 2s
each = up to 200s) on a nightly cadence while keeping the
fast-finding tests (theory matrix of bit-flip seeds, etc.) in the
default unit-test run. Lower friction than a separate solution and
CI workflow.

**Deterministic seeds for the bit-flip theory.** `Random(seed)` with
literal-int seeds in `[InlineData]` makes any future failure
reproducible by re-running that exact test case. The bulk sweep uses
`new Random(42)` for the same reason — sweep is deterministic across
runs, so a new finding bisects cleanly to a specific iteration index.

### What the harness will likely find next

Educated guesses for future findings (each is a one-line addition to
the filter when it surfaces):

- Specific NPOI-internal exceptions that aren't yet in the
  `"NPOI."` prefix match (some are nested types with weird
  full-name shapes).
- `ICSharpCode.SharpZipLib.Zip.ZipException` — currently caught via
  the namespace prefix, but specific subclasses might have surprising
  shapes.
- `System.Threading.ThreadAbortException` if a thread-pool worker
  gets canned mid-parse.

When any of these show up, the fix is the same shape: extend
`IsKnownMalformedOpenException`, add a regression test, write a
note here.

---

## Future entries

Add a dated section per substantive implementation milestone. After
v1.0 ships, distill the patterns that recurred into a sibling methodology
document under `~/dev/projects/references/project-methodologies/`.
