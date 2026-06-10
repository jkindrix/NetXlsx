# S2 design-proposal memo — 2026-06-10

**Status: AWAITING OPERATOR SIGN-OFF.** No code lands from this memo until each
proposal is individually signed off (remediation-ledger rule 2: propose the
shape, do not implement in the same session). On sign-off, each I-NN below is
copied into `design.md` §6 in house style by the slice that implements it, and
the corresponding R-row graduates.

Contents:

- [I-88](#i-88) — control-character policy (R-3) → lands S10 first
- [I-89](#i-89) — theme cluster (R-8 + XlsxCodeGen Appendix A #1) → lands second
- [I-90](#i-90) — sheet lifecycle: rename / move / delete (R-12) → lands third
- [I-91](#i-91) — removal family (R-11, folds in R-10) → lands fourth, 2 slices
- [Operator decisions R-31..R-34](#operator-decisions) — one memo, recommendations included

Numbering: design.md's highest decision is I-87 (bulk-write row-index cache),
so these take I-88..I-91, numbered in their S10+ landing order.

---

## I-88 — Control characters: lossless `_xHHHH_` escaping for text, fail-fast for names/formulas (R-3) {#i-88}

### Problem

XML-1.0-invalid characters (0x00–0x08, 0x0B, 0x0C, 0x0E–0x1F, U+FFFE/FFFF,
lone surrogate halves) in `SetString`, rich text, or `Comment` pass the setter
silently and explode at `Save` as a raw `ArgumentException` from the SDK — no
cell context, outside the `WorkbookException` family, and (pre-S1) combined
with the truncate-first save into data loss. Verified live (ledger R-3).
ECMA-376 defines `ST_Xstring` (§22.9.2.19) with the `_xHHHH_` escape
convention exactly for this; Excel itself writes these escapes. The
competitive field: the raw SDK throws, SpreadCheetah silently **strips**
(lossy — verified in its `XmlUtility.cs`). A lossless implementation exceeds
every library in the set.

### Proposed shape

No new public API. Behavioral decision in two halves:

**(b) Lossless escaping — user *content* surfaces.** Applies to `SetString`,
`SetRichText` (per-run), `Comment`. At the setter, encode to the `ST_Xstring`
convention before handing text to the SDK:

- Each XML-invalid char → `_xHHHH_` (four uppercase hex digits of the UTF-16
  code unit; lone surrogate halves escape as their own code unit, e.g.
  `_xD800_`).
- Any *literal* substring in user text that matches the escape pattern
  (`_xHHHH_`) gets its leading underscore escaped: `_x005F_xHHHH_` — required
  by the convention for round-trip fidelity.
- Valid-XML whitespace (tab 0x09, LF 0x0A, CR 0x0D) passes through untouched
  (already round-trips fine — pinned by existing tests).

**Read-side decode (the other half of lossless).** `GetString`, `GetComment`,
rich-text read, and therefore generated `ReadRows`, decode well-formed
`_xHHHH_` sequences (exactly 4 hex digits) back to the original characters;
`_x005F_` decodes to a literal underscore that re-protects the following
sequence. This is also a latent read-fidelity fix: today an *Excel-authored*
file containing `_x0008_` reads back as the seven-character literal, not the
control char Excel meant. Both engines store the escaped form in the DOM, so
the package bytes are exactly what Excel would write. Note for the slice:
`TextForMeasurement` (AutoSize) must measure the *decoded* form.

**(a) Fail-fast — *name/formula* surfaces**, where control chars are never
meaningful and escaping would corrupt semantics:

- Sheet names: extend `ValidateSheetName` / `IsValidSheetName` /
  `SanitizeSheetName` (the shared char array — same mechanism R-9's S6 slice
  tightens) to reject XML-invalid chars → `SheetNameException` naming the
  offending char.
- Defined names: `ValidateDefinedName` already restricts the charset; verify
  it excludes the control range, and validate the *formula body* (currently
  unvalidated — overlaps R-16's named-range item) at least for XML-invalid
  chars → existing validation exception.
- Formulas: `SetFormula`'s structural validation additionally rejects
  XML-invalid chars → `FormulaException` naming the cell.

All failures are at the **setter**, in the `WorkbookException` family, naming
the location — never a raw save-time `ArgumentException` (decision #21,
fail loud, fix early).

Both engines: the streaming writer escapes identically in its cell/text
emission path.

### What this deliberately does not do

- No `WorkbookOptions` knob to choose strip/throw instead: escaping is
  lossless and Excel-native, so there is nothing to opt out *to*. A knob can
  be added compatibly later if real demand appears.
- No escaping in sheet names/defined names/formulas (Excel rejects control
  chars there at entry; escaping would change reference semantics).

### Tests / verification

- Round-trip every char in 0x00–0x1F (escaped vs passthrough partition),
  U+FFFE/FFFF, lone surrogate halves, and literal `_x0008_` input
  (`_x005F_`-protected) — through `SetString`, rich text, comments, both
  engines, plus a `ReadRows` compile-and-run path.
- Excel-authored fixture via the crafted-cell zip-injection pattern
  (`x:`-prefixed, per the review's reusable-assets note) containing real
  `_xHHHH_` escapes → `GetString` decodes.
- Oracles: LO opens the output and displays/preserves the text; pin
  openpyxl/calamine actual behavior in-slice (do not assume they decode).
- Fail-fast probes for names/formulas (exception type + message names the
  offender).
- Golden impact: **none** — no existing golden can contain these chars (the
  writer used to throw), and escaping changes no other byte.

Closes R-30's control-char/lone-surrogate test gap as a side effect.

---

## I-89 — Theme cluster: lazy default-theme embed + theme-color styling symmetry (R-8 + XlsxCodeGen Appendix A #1) {#i-89}

### Problem

`SetThemeXml` (I-79), `GetThemeXml`/`ResolveThemeColor` (I-81), and
`CellStyle.BackgroundTheme` ship today, but `Create()` embeds **no theme
part** while theme-indexed styling is accepted — so theme-indexed colors are
a consumer lottery (verified: LO injects its *own* theme — theme-4 resolves
`#FF18A303` green vs Office blue — and rewrote a theme-indexed picture border
to literal white). Separately, Appendix A #1 (the largest gap XlsxCodeGen's
scout found): font/run/border theme colors have **no write surface and read
back `null`** — only `BackgroundTheme` exists.

### Options considered

- **(a-eager) Embed a default Office theme in every `Create()` output.**
  Matches Excel's own behavior, but adds ~10 KB to every minimal workbook,
  regenerates *all* goldens, and pays the cost even for the majority of
  workbooks that never touch a `ThemeColor`. Against the thinness principle.
- **(b) Fail loud on `ThemeColor` styling without a theme part.** Rejected:
  retroactively breaks the shipped, documented `BackgroundTheme`/I-86 surface
  and turns a fixable output gap into a user-facing chore.
- **(c) Ship a `DefaultTheme` constant + document the footgun.** Cheapest,
  but leaves the default path broken-by-default; docs are not a fix.
- **(a-lazy) — PROPOSED.** Embed the default theme automatically at the
  moment it becomes load-bearing: the first time a theme-indexed color is
  *written* (style registration, picture border, …) into a workbook that has
  no theme part. Plain workbooks stay byte-identical; theme-using workbooks
  become self-contained the way Excel's are.

### Proposed shape

```csharp
// On the static Workbook facade:
public static byte[] DefaultThemeXml { get; }   // fresh copy per call; the standard
                                                // Office theme, hand-authored XML
                                                // (values are public facts), suitable
                                                // for SetThemeXml(Workbook.DefaultThemeXml)

// Styling symmetry (Appendix A #1) — mirrors shipped BackgroundTheme exactly:
public sealed record CellStyle { ..., ThemeColor? FontColorTheme { get; init; } }
public sealed record RichTextStyle { ..., ThemeColor? ColorTheme { get; init; } }
public sealed record CellBorders { ..., ThemeColor? TopColorTheme,
                                        ThemeColor? BottomColorTheme,
                                        ThemeColor? LeftColorTheme,
                                        ThemeColor? RightColorTheme { get; init; } }
```

Behavior:

- **Lazy embed:** when a `ThemeColor`-carrying style element (background,
  font, border, rich-text run, picture border, connector scheme color) is
  written and `GetThemeXml()` is null, the workbook embeds
  `DefaultThemeXml` first. Explicit `SetThemeXml` before or after still wins
  (replaces, invalidates the I-81 cache as today). Opened third-party files
  without a theme get the embed only if *new* theme styling is written.
- **Precedence rule** (per slot, both CellStyle and borders): theme variant
  wins over the literal-color property when both are set — the rule
  `BackgroundTheme` already established; restate it on every new property's
  XML docs.
- **Read-back:** the style read path populates the new theme properties from
  `theme`/`tint` attributes (today they read back null — Appendix A #1's
  exact complaint), leaving literal-color properties null for theme-indexed
  XML, and vice versa.
- **I-81 contract unchanged:** `ResolveThemeColor` still returns null when no
  theme part exists; after the first theme-styling write it resolves against
  the embedded default — the documented behavior simply starts holding on
  fresh output.

### Tests / verification

- Fresh workbook + `FontColorTheme` → theme part present, openpyxl reads
  `theme`/`tint` on the font, `ResolveThemeColor(4)` non-null and equal to
  the Office accent value.
- LO round-trip: theme-indexed border/font survives resave with *Office*
  colors (the lottery probe inverted).
- Plain workbook (no theme styling) → byte-identical to today (no theme part).
- Read-back fixtures (Excel/LO-authored) for font/run/border theme+tint.
- Golden impact: **[golden-regen]** for goldens already using `ThemeColor`
  (the I-86 picture-border goldens at minimum) — deliberate, called out in
  the slice commit. Plain goldens untouched.

Converges with R-33's "full theme-color styling" triage row — signing this
off answers that row's *shape*; R-18's re-baseline schedules it.

---

## I-90 — Sheet lifecycle: `Rename` / `MoveSheet` / `RemoveSheet` (R-12) {#i-90}

### Problem

No sheet rename, reorder, or delete (`ISheet.Name` get-only; nothing removes
or moves a sheet) — basic capability every competitor has, and the roadmap
matrix falsely marks it shipped in v1.0 (R-17/R-18 correct the matrix
regardless of this proposal's fate). The internal seam already anticipates
it: `OoxmlSheet.SetNameInternal` carries a "if a rename API lands" comment.

### Proposed shape

```csharp
// ISheet
void Rename(string newName);          // SheetNameException on invalid/duplicate
                                      // (AddSheet rules, incl. R-9's tightened set)

// IWorkbook
void MoveSheet(ISheet sheet, int newIndex);   // 1-based resulting position
                                              // (remove-then-insert; matches
                                              // AddSheet(name, index) semantics)
void RemoveSheet(ISheet sheet);
```

**Rename is a method, not a `Name` setter**: rename validates, can throw
`SheetNameException`, and rewrites references document-wide — side effects a
property setter would hide. (Rejected: `Name { set; }`.)

**Reference rewrite on rename — included, via a sheet-reference lexer, not a
formula parser.** The engine has no formula parser, but rewriting sheet
references only requires recognizing `'Quoted Name'!` and `BareName!` prefixes
*outside string literals* — and `SetFormula`'s structural validator already
lexes quoted-sheet-name and string literals. The lexer rewrites the old name
(case-insensitive, per sheet-name semantics) in:

- cell formulas (`<f>`, including shared-formula masters) on **all** sheets,
- defined-name bodies,
- internal hyperlink locations (`#Sheet!A1`),
- conditional-formatting and data-validation formulas,
- chart series references (`c:f`).

Quoting is normalized on output (quote the new name iff it needs quoting;
double embedded apostrophes). Documented residual, which is **Excel parity**:
string arguments (e.g. `INDIRECT("Old!A1")`) are not rewritten — Excel
doesn't rewrite those either.

**Delete (`RemoveSheet`)** — follows the `RemoveTable` precedent
(`ArgumentException` on a foreign or stale handle) plus:

- `InvalidOperationException` when it would leave zero visible sheets (a
  valid workbook needs at least one).
- Part cleanup: the `WorksheetPart` and its owned descendants (drawings,
  comments, tables) via `DeletePart`; the `<sheet>` entry and its
  relationship; the calcChain part if present (we never write one, but opened
  files may carry one — Excel rebuilds it).
- Defined names scoped to the deleted sheet are removed; `localSheetId` on
  every later-sheet-scoped name is re-indexed (it is a zero-based sheet
  *position*).
- Cross-sheet references to the deleted sheet rewrite to `#REF!` via the same
  lexer — Excel parity, and honest where silence would corrupt.
- `bookViews/@activeTab` clamped into range.
- The removed sheet's wrapper (and its cells') members thereafter throw
  `InvalidOperationException` ("sheet has been removed") — distinct from
  `ObjectDisposedException`, matching the design's existing
  removed-sheet-access language.

**Move (`MoveSheet`)** — reorders `<sheets>` and the index list; re-indexes
`localSheetId` on all sheet-scoped defined names and clamps/follows
`activeTab`. The workbook's positional indexer and `Sheets` order reflect the move. Throws
`ArgumentOutOfRangeException` for an out-of-range index, `ArgumentException`
for foreign/stale handles.

**Streaming engine: out of scope** — `IStreamingWorkbook` is forward-only;
these members are simply not added there.

### Slicing note

This is the largest of the four. If the operator prefers, it splits cleanly:
rename+move (lexer, re-indexing) in one slice, delete (part cleanup, `#REF!`)
in the next — the lexer is shared and lands with the first.

### Tests / verification

- Rename: collision/invalid-name throws; formulas, defined names, hyperlinks,
  CF/DV, chart refs rewritten across sheets (quoted and unquoted, names
  needing quoting after rename); `INDIRECT` string untouched; LO + openpyxl
  oracles confirm references still resolve.
- Delete: last-sheet guard; part listing shows no orphaned parts or rels
  (zip-level assert); `#REF!` rewrite; `localSheetId` re-index; stale-handle
  throws; openpyxl opens the result cleanly.
- Move: index semantics pinned at both ends and middle; `localSheetId` +
  `activeTab` follow; round-trip order preserved in LO.
- Golden impact: none (new API; goldens don't exercise it).

---

## I-91 — Removal family: symmetric removal for hyperlinks, comments, drawings, names, validations (R-11, folds in R-10) {#i-91}

### Problem

Removal asymmetry: `UnmergeCells` / `Unprotect` / `RemoveTable` /
`RemoveConditionalFormatting` / `ClearAutoFilter` exist; nothing removes
hyperlinks, comments, pictures, charts, shapes, connectors, named ranges, or
validations. R-10 (verified): `SetString` after `Hyperlink` keeps the rel
pointing at the old URL under new text, and there is no way off.

### Naming rule (codified)

The existing surface implies it; this row makes it law: **`Remove<Thing>`**
for discrete added objects (keyed by handle or unique key), **`Clear*`** for
singleton sheet state, **`Un*`** only for the two shipped legacy verbs
(merge/protect — grandfathered, not extended).

### Proposed shape

```csharp
// ICell — fluent, idempotent (no-op when absent), returns ICell like
// Hyperlink()/Comment() do:
ICell RemoveHyperlink();
ICell RemoveComment();

// ISheet — handle-based, RemoveTable semantics exactly
// (ArgumentException on foreign or stale handle):
void RemovePicture(IPicture picture);
void RemoveChart(IChart chart);
void RemoveShape(IShape shape);
void RemoveConnector(IConnector connector);

// ISheet — key-based, UnmergeCells semantics (exact-range match required;
// InvalidOperationException when no validation has that exact range):
void RemoveValidation(string a1Range);

// IWorkbook — key-based (names are unique, case-insensitive per I-9);
// ArgumentException when absent, matching RemoveTable's not-found behavior:
void RemoveNamedRange(string name);
```

Cleanup contracts (the part/rel discipline `RemoveTable` and the
no-empty-artifact comment rules already establish):

- `RemoveHyperlink`: deletes the `<hyperlink>` element and, for external
  targets, the reference relationship (the code path `SetHyperlink`'s replace
  logic already has); drops the `<hyperlinks>` container when empty.
- `RemoveComment`: removes the comment + author bookkeeping from the comments
  part and the VML popup shape; deletes the comments part and VML part
  entirely when the last comment goes (no empty zero-index artifact —
  existing discipline).
- Drawing-layer removals: delete the anchor from the drawing part; delete the
  image/chart part **only when no other anchor references it** (the slice
  verifies whether `AddPicture` dedups image parts and guards accordingly);
  drop the drawing part + worksheet rel when the last anchor goes.
- `RemoveNamedRange`: deletes the `<definedName>`; drops `<definedNames>`
  when empty.
- `RemoveValidation`: deletes the `<dataValidation>`; drops the container +
  `@count` bookkeeping when empty.
- After a handle-based removal, the handle's members throw
  `InvalidOperationException` (stale-handle pattern, as `RemoveTable`).

**R-10 fold-in:** `ICell.Hyperlink`/`SetString` XML docs state the
keep-link-on-text-edit semantics (Excel parity) and point at
`RemoveHyperlink`. The defect was the missing exit + missing docs; both close
here.

### Slicing

Two slices after sign-off, per the ledger's estimate:

1. **Cell/workbook level:** `RemoveHyperlink`, `RemoveComment`,
   `RemoveNamedRange`, `RemoveValidation` (element + rel cleanup, simple).
2. **Drawing layer:** pictures/charts/shapes/connectors (anchor + shared-part
   reference counting — the riskier half).

### Tests / verification

- Each removal: element gone, part/rel gone when last (zip-level part-listing
  asserts), container dropped when empty, no-op vs throw semantics pinned per
  the table above, stale-handle throws.
- R-10 regression: `Hyperlink` → `RemoveHyperlink` → `SetString` → openpyxl
  confirms no rel, no `<hyperlink>`.
- Shared image part: two pictures, remove one → other still renders (LO
  oracle), part survives; remove both → part gone.
- Golden impact: none.

---

## Operator decisions R-31..R-34 (recommendations attached) {#operator-decisions}

### R-31 — Claim the `NetXlsx` nuget.org ID?

The ID is unclaimed (404, verified 2026-06-10). Squatting is unguarded until
the first push; the name has invested identity (repo, docs, generated code,
diagnostics prefix). Publication proper stays operator-gated — this is only
about reserving the name.

| Option | Notes |
|--------|-------|
| **Push a `0.0.1` placeholder (RECOMMENDED)** | Cheap, immediate, fully under our control. Package description says "placeholder — first real release pending"; can be unlisted immediately after push so it never surfaces in search, while still owning the ID. |
| Apply for `NetXlsx.*` prefix reservation | Stronger (covers satellite packages) but requires a Microsoft review process and is *in addition to*, not instead of, owning the bare ID. Can follow later. |
| Do nothing until first real publish | Free, but leaves the name exposed for however long S2→S10+ takes. The risk is small in absolute terms but the mitigation costs minutes. |

**Recommendation:** reserve now via unlisted `0.0.1` placeholder; consider
prefix reservation at first real publish. Needs only: a nuget.org account
decision and one manual push (operator action — not a CI change).

### R-32 — Release workflow modernization at publish time

**Recommendation: adopt both, but only when R-31/publication becomes real**
(no code now): NuGet Trusted Publishing (OIDC via `NuGet/login@v1`) instead
of a long-lived `NUGET_API_KEY` secret, and
`actions/attest-build-provenance` on the GH-release artifacts — with the
documented caveat that nuget.org re-signs nupkgs, so attestation covers the
GH artifacts, not the gallery copy. Folds into the S7 infra slice **only if**
R-31 lands a placeholder (Trusted Publishing is configured per-package on
nuget.org, so it needs the ID to exist); otherwise it waits with publication.

### R-33 — Post-swap feature triage (the dead-NPOI-blocker list)

These all carried "Blocked on NPOI 3.x" rationales for an engine that no
longer exists. Verdicts feed R-18's matrix re-baseline:

| Feature | Recommendation | Rationale |
|---------|---------------|-----------|
| Full theme-color styling | **Schedule — shape is I-89** | Already designed above; pairs with R-8's defect fix. |
| Streaming read (`OpenXmlPartReader`) | **Promote (v2.x)** | SDK has the primitive; completes the streaming story the README sells; pairs with R-34's `IAsyncEnumerable` ask. |
| `In(...)` 3+ values | **Promote (small)** | Probe-verified the gap is model-shape-only on the SDK engine — the `NotSupportedException` guards a limitation that died with NPOI. Cheap, removes a documented wart. |
| AutoFilter Top-N | **Schedule (v2.x, behind the two above)** | SDK can author `<top10>` directly; demand is thinner than streaming read. |
| Threaded comments | **Demote (hold for demand)** | Legacy comments round-trip fine; threaded adds parts + person registry for a niche ask. Revisit if a user asks. |

### R-34 — Strategic feature bets (roadmap material)

**Recommendation: decide these at the R-18 re-baseline session (S5), with the
following defaults**, demand-ranked from the ecosystem research:

1. **"Coming from EPPlus/ClosedXML" migration guide — do first, it's a docs
   slice.** EPPlus-8 license keys + ClosedXML's 13-month stable gap is an
   active migration window; this is the cheapest capture of it.
2. **Pivot-table authoring — confirm as the headline v3.0 bet.** Competitors
   are chronically broken here; the preservation-first architecture is
   genuinely advantaged.
3. **Agile encryption (MS-OFFCRYPTO) — v3.0, evaluate as an add-on package**
   so the core stays thin/AOT-clean.
4. **`IAsyncEnumerable` async streaming — bundle with R-33's streaming-read
   slice**, not standalone.
5. **AutoSize approximate-metrics fallback — small, schedule
   opportunistically** (softens I-84's `MissingFontException` edge).

---

## Sign-off checklist

- [ ] I-88 control-character policy (escape content / fail-fast names) — approve / amend
- [ ] I-89 theme cluster (lazy default embed + styling symmetry) — approve / amend; eager-embed alternative explicitly rejected?
- [ ] I-90 sheet lifecycle (method rename + lexer rewrite; delete; move) — approve / amend; one slice or split?
- [ ] I-91 removal family (naming rule + surface + 2-slice plan) — approve / amend
- [ ] R-31 nuget.org ID — reserve via placeholder? (operator action)
- [ ] R-32 trusted publishing + attestation — adopt-at-publish?
- [ ] R-33 triage verdicts — confirm for the R-18 re-baseline
- [ ] R-34 strategic defaults — confirm for the R-18 re-baseline
