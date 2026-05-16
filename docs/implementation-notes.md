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
- To "clear" a style entirely, pass `CellStyle.Default` — but that
  clears EVERYTHING via the merge, because Default is all-null and
  the existing cell's properties have nothing to merge against
  (current → empty → no axes). Actually no, the merge is
  `existing ← overlay-non-null`, so Default doesn't clear anything;
  it's a no-op. Real clearing needs an explicit sentinel which we
  haven't added yet. Documenting as a known gap; minor.

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

---

## Future entries

Add a dated section per substantive implementation milestone. After
v1.0 ships, distill the patterns that recurred into a sibling methodology
document under `~/dev/projects/references/project-methodologies/`.
