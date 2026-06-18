# Design proposal — I-92: open workbooks containing chartsheets / dialogsheets (R-38)

**Status: PROPOSED — awaiting operator sign-off.** Per the ledger's
`[I-NN first]` rule, this memo proposes the shape only; no code lands in the
same session. On sign-off it graduates R-38, is copied into `design.md` §6 in
house style by the implementing slice (S20), and the R-38 row closes.

**Numbering:** the highest decision today is I-91 (removal family); this is
**I-92**.

---

## Problem

A *valid*, Excel-authorable `.xlsx` whose `<sheet>` entry targets a
`ChartsheetPart` (a full-window chart, no cell grid) — or the legacy
`DialogsheetPart` — **cannot be opened**. `OoxmlWorkbook.IndexExistingSheets`
resolves each `<sheet>`'s `r:id` to a part and casts to `WorksheetPart`; a
non-worksheet part falls through to a `MalformedFileException`
(`src/NetXlsx/Internal/OoxmlWorkbook.cs:340-343`):

```csharp
if (part is not WorksheetPart wsPart)
    throw new MalformedFileException(
        $"... targets a {part.GetType().Name}, not a worksheet part" +
        " (chartsheet/dialogsheet workbooks are not supported yet — tracked as ledger R-38)");
```

This is honest (post-R-37 it is a classified `MalformedFileException`, not the
pre-R-37 `InvalidCastException`) but **wrong for a legal file**. Chartsheets
appear in the workbook tab order like any sheet; a library that opens
arbitrary user-supplied workbooks will eventually meet one and hard-reject the
whole file. For a fidelity-first library this is the wart that erodes trust,
and the §7.7 unknown-part-preservation promise is currently moot for these
files because the *open itself* fails before preservation can apply.

**Scope of the artifacts.** Chartsheets are a first-class, current Excel
feature ("Move Chart → New sheet"): uncommon but not rare in real workbooks.
Dialogsheets are an extinct Excel-5.0/95 macro-dialog construct — they still
have a legal `<dialogsheet>` element and `DialogsheetPart`, but no modern
authoring path. Both should *open*; only chartsheets warrant any thought
beyond "don't reject."

---

## Options (recap, with the decisive fact)

**Decisive fact:** the chartsheet part already round-trips today. It is
reachable in the OPC relationship graph (that is how `GetPartById` resolves
it), so the clone-based `Save` (`_document.Clone`) preserves it verbatim — it
is not even an orphan needing the `CaptureOrphanParts`/`ReinjectOrphanParts`
machinery. **The only thing between today's hard reject and a clean
open+round-trip is the line-340 cast.** That is what makes Option A small.

- **Option A — skip-with-visibility (RECOMMENDED).** The chartsheet appears in
  `SheetCount`, the names list, `Sheets`, and the indexers, represented by a
  minimal placeholder. Any cell-/grid-shaped access throws a documented
  exception. The part is preserved untouched on save. Small, near-zero public
  surface, forward-compatible with B.
- **Option B — full `IChartsheet` wrapper.** Model and read the chartsheet's
  chart (name, the chart itself, page setup, tab color). Large, new permanent
  public type, ongoing maintenance and test surface — not justified by how
  rare chartsheets are, and most consumers only need the file to *open and
  round-trip*, not to manipulate the embedded chart programmatically.
- **Option C — defer.** Keep rejecting. Cheapest now; leaves NetXlsx
  hard-rejecting legal Excel files. Rejected: the fix is cheap and the wart is
  real.

**Recommendation: Option A**, forward-compatible with B should demand for
programmatic chart access ever appear.

---

## Proposed shape (I-92, Option A)

### Public surface — minimal

One new discriminator on `ISheet` so callers can detect a non-grid sheet
before touching it:

```csharp
// On ISheet:
SheetKind Kind { get; }       // Worksheet (default) | Chartsheet | Dialogsheet

public enum SheetKind { Worksheet, Chartsheet, Dialogsheet }
```

Rationale for an **enum over a `bool IsChartsheet`**: a single bool would have
to read `true` for dialogsheets too (misleading), or we'd need a second bool.
`Kind` is honest, one member, and future-proof (a later Option-B
`IChartsheet` can hang off the same discriminator). *(Open question 1 — see
below — if you prefer the smaller `bool IsChartsheet`, say so.)*

No other new public members. The chartsheet placeholder is still an `ISheet`.

### Behavior on Open

Replace the line-340 throw with a three-way classification:

1. `WorksheetPart` → today's path (wrap in `OoxmlSheet`, `NormalizeMissingReferences`).
2. `ChartsheetPart` / `DialogsheetPart` → construct a **placeholder sheet**
   recording: name, the backing part, tab position, `Hidden` state (from
   `<sheet state>`), and `Kind`. It participates in `SheetCount`, `Sheets`,
   `this[int]`, `this[string]`, `TryGetSheet`. **`NormalizeMissingReferences`
   (worksheet-grid-specific) is skipped** — the part is never rewritten, so
   round-trip stays byte-stable.
3. anything else (e.g. a sheet `r:id` mis-pointed at the styles part) →
   `MalformedFileException`, **unchanged**. This keeps
   `MalformedInputContractTests.Sheet_Targeting_NonWorksheet_Part_Fails_Loud`
   green — only legitimately *sheet-typed* parts get the placeholder.

### What works vs. what throws on a chartsheet placeholder

**Works** (workbook-level / non-grid):
- `Name` (get), `Kind`, `Hidden` (get/set — visibility is a `<sheet>`
  attribute, not grid).
- `Rename` — renaming a chartsheet is legal; it is a `<sheet name>` change +
  the same document-wide reference rewrite, no grid involved. Keeping it works
  preserves I-90 lifecycle parity.

**Throws** a documented exception — every grid-/content-shaped member:
indexers (`this[a1]`, `this[row,col]`), `Range`, `AppendRow`, `Row`, `Column`,
`LastRowNumber`, `Freeze*`, `Group*`/`Ungroup*`, `CreateSplitPane`,
`AddChart`/`AddShape`/`AddPicture`/`AddConnector` (+ their `Remove*`),
`AddConditionalFormatting`/`ConditionalFormattingCount`/`Remove…`, `SortRange`,
`MergeCells*`/`UnmergeCells`, `ShowGridlines`, `AddTable`/`Tables`/`TryGetTable`/
`RemoveTable`, `SetAutoFilter*`/`ClearAutoFilter*`/`HasAutoFilter`/
`AutoFilterRange`, `AddValidation`/`RemoveValidation`.

**Exception type** *(open question 2)*: recommend a plain
`NotSupportedException` with a recoverable message —

> *"Sheet 'Chart1' is a chartsheet and has no cell grid; cell/grid access is
> not supported. Use IWorkbook.Underlying to reach the ChartsheetPart."*

`NotSupportedException` is the idiomatic .NET signal for "this member is not
valid on this object" and needs no new type. The alternative is a new
`ChartsheetException : WorkbookException` for family consistency/greppability;
I lean to `NotSupportedException` for the smaller surface.

### Lifecycle interaction (I-90)

- **`MoveSheet`** — works unchanged (operates on `_sheetsByIndex` order and the
  `<sheets>` element; the placeholder participates naturally).
- **`RemoveSheet`** — needs a small branch. Today it resolves
  `target.WorksheetPartInternal` and `DeletePart`s it; for a chartsheet it must
  resolve and `DeletePart` the `ChartsheetPart`/`DialogsheetPart` instead (the
  Clone-reachability teardown of sub-parts is identical). The last-VISIBLE-sheet
  guard still applies (a non-hidden chartsheet counts as visible). The
  worksheet-only cleanups (calcChain, pivot caches sourced from the sheet) are
  no-ops for a chartsheet; the scoped-defined-name purge stays (a chartsheet can
  be a name scope).
- **`activeTab`** clamp/follow logic is positional and already handles a
  chartsheet at any index.

### Round-trip preservation

The placeholder must never rewrite its part. Open → Save keeps the
`ChartsheetPart` byte-stable (proven mechanism: §7.7 + the existing
clone-Save). This is asserted directly in the test plan.

---

## Test plan

A new `ChartsheetTests` (engine-level), driven by an **Excel-/LibreOffice-
authored fixture** (NetXlsx cannot author chartsheets): one workbook with a
worksheet **and** a chartsheet in tab order.

1. **Open** — `SheetCount` includes the chartsheet; `this[name]`/`Sheets`
   return it; `Kind == Chartsheet`; the worksheet beside it reads normally.
2. **Grid access throws** — a representative set of grid members throw the
   documented exception with the sheet name in the message.
3. **Round-trip** — open → save → the `ChartsheetPart` (and its chart sub-part)
   survive; reopen still works; a zip-level no-orphan/no-dangling-rel assert
   passes.
4. **Lifecycle** — `Rename` the chartsheet (name changes; references rewrite);
   `MoveSheet` reorders it; `RemoveSheet` drops it and its part (no-orphan
   assert), and the last-visible guard fires when removal would leave no
   visible sheet.
5. **`Hidden`** get/set on the chartsheet.
6. **Regression** — `Sheet_Targeting_NonWorksheet_Part_Fails_Loud`
   (styles-part misdirection) still throws `MalformedFileException`.
7. **Interop (optional, if cheap)** — LO 26.2 resave of the round-tripped file
   still renders the chartsheet.

Dialogsheet coverage rides the same placeholder branch; a dedicated fixture is
optional given extinction (synthesize via `Underlying`, or skip and note it).

**PublicApi:** `+ NetXlsx.SheetKind` (enum + members) and
`NetXlsx.ISheet.Kind.get` in `PublicAPI.Unshipped.txt` (or one bool if you pick
that). **Effort:** small-to-medium, one slice.

---

## Open questions for the operator

1. **Discriminator shape** — `ISheet.Kind` enum *(recommended)* vs
   `bool ISheet.IsChartsheet`.
2. **Grid-access exception** — `NotSupportedException` *(recommended, no new
   type)* vs a new `ChartsheetException : WorkbookException`.
3. **Lifecycle on a chartsheet** — `Rename`/`MoveSheet`/`RemoveSheet` work
   *(recommended — parity)* vs treat a chartsheet as fully opaque (those throw
   too).
4. **Dialogsheets** — same placeholder, `Kind == Dialogsheet` *(recommended)*;
   no separate modelling.

---

## Sign-off checklist

- [ ] Option A approved (skip-with-visibility), B/C rejected.
- [ ] Discriminator shape chosen (Q1).
- [ ] Grid-access exception chosen (Q2).
- [ ] Lifecycle disposition chosen (Q3).
- [ ] Dialogsheet disposition chosen (Q4).
- [ ] On sign-off: graduate R-38, implement as slice S20, copy into design.md §6
      (next free §6.12.x), add the Excel/LO chartsheet fixture.
