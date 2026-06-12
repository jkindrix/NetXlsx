# Interop & limits

**Status:** ACTIVE — established 2026-06-12 (ledger R-22). Every claim here
is evidence-backed: the LibreOffice matrix was probe-verified during the
2026-06-10 full-repo review and is now **asserted nightly** by
`.github/workflows/interop-nightly.yml` (LibreOffice headless resave +
openpyxl read-only, over the kitchen-sink + streaming workbooks generated
by `tools/NetXlsx.InteropProbe`). When the nightly job and this document
disagree, the job is the truth and this file has drifted — fix it.

## LibreOffice resave matrix

What survives a LibreOffice Calc open → save-as-xlsx cycle of a
NetXlsx-authored workbook (LO 26.2.4 at verification; the nightly tracks
the current `libreoffice-calc` package):

| Surface | LO resave | Notes |
|---|---|---|
| Cell values (string/number/bool/date/time/duration) | **survives** | LO normalizes single-cell ranges in formulas (`SUM(A2:A2)` → `SUM(A2)`) |
| Formulas + named ranges (workbook & sheet scope) | **survives** | LO calculates on load and writes cached results — NetXlsx typed getters read them (R-7) |
| Merges, rich text, hidden rows/cols/sheets | **survives** | rich-text runs may be restructured (inline → shared strings); formatting kept |
| Tables + totals rows | **survives** | |
| Conditional formatting | **survives** | |
| Comments, hyperlinks (incl. `mailto:`) | **survives** | targets round-trip verbatim (SDK-quirk #15) |
| Pictures, shapes, connectors | **survives** | anchors kept |
| Sheet protection (incl. password verifier) | **survives** | |
| Charts | renders | re-emitted in LO's own chart XML |
| Sheet `<autoFilter>` + criteria | **DROPPED** | LO keeps the filtered *values*, discards the filter model |
| Workbook structure lock (`Protect`) | **DROPPED** | sheet-level protection is kept; the workbook-level lock is not |
| Theme part | **REPLACED** | LO substitutes its own theme: theme-indexed colors resolve differently (theme-4: Office blue → LO green `#FF18A303`), and theme-indexed picture borders are rewritten to literal colors (observed: white). Until I-89 lands (default theme at `Create()`), theme-indexed styling is consumer-dependent — prefer literal colors when LO is in the consumer set |

## Hyperlink-on-edit semantics

`SetString` on a cell that carries a hyperlink **keeps the hyperlink**
pointing at its original target under the new text (Excel behaves the same
way on text edit). There is currently no API to remove a hyperlink — the
removal family is design I-91 (signed off 2026-06-11) and lands per the
remediation ledger (R-10/R-11). Until then: to "remove" a link, reach
through `ISheet.Underlying` and delete the `<hyperlink>` element plus its
package relationship.

## Size ceiling

`System.IO.Packaging` (under the Open XML SDK) holds a **~2 GB
single-part ceiling** (the SDK's top-voted issue). For NetXlsx this is a
*per-part* limit — in practice the worksheet XML of one enormous sheet —
and it applies to the streaming writer too: streaming bounds *memory*, not
part size. Workloads approaching it should shard across sheets or files.

## Formula injection (untrusted data)

NetXlsx writes what you pass. If cell text originates from untrusted
input, a value beginning with `=`, `+`, `-`, or `@` becomes a live
formula in consumers that evaluate (Excel, LO) — the classic CSV/DDE
injection vector (`=WEBSERVICE(...)`, `=cmd|...` style payloads).
NetXlsx deliberately does not second-guess values (`SetString` stores a
string cell, which is inert — the vector is callers building *formulas*
from user input, or downstream CSV exports). Guidance: never interpolate
untrusted input into `SetFormula`/named-range formulas; for display-only
data always use `SetString` (inert by construction); if you re-export to
CSV, apply the usual leading-character escaping there.

## `docProps` (core/app metadata)

NetXlsx-authored packages carry **no `docProps/core.xml` or
`docProps/app.xml`** ([verified] 2026-06-10 by part-list probe). This is
legal OPC and every consumer in the gauntlet opens the files fine; it is
merely unconventional (most producers stamp creator/created). **Decision
(R-22, 2026-06-12): stay as-is.** Rationale: metadata is consumer-optional;
emitting it would add non-deterministic bytes (timestamps) that the
golden/fixture normalization would immediately have to strip back out; and
callers who want metadata can add the parts through
`IWorkbook.Underlying`. Revisit on user demand via a ledger entry.

## Source-generator flow for source-tree consumers

NuGet consumers get the `[Worksheet]` generator automatically
(`analyzers/dotnet/cs/` in the package). **A bare `<ProjectReference>` to
`src/NetXlsx/NetXlsx.csproj` does NOT flow the generator** — source-tree
consumers must also reference the generator project explicitly:

```xml
<ProjectReference Include="…/src/NetXlsx/NetXlsx.csproj" />
<ProjectReference Include="…/src/NetXlsx.SourceGen/NetXlsx.SourceGen.csproj"
                  OutputItemType="Analyzer"
                  ReferenceOutputAssembly="false" />
```

(`tools/NetXlsx.InteropProbe/NetXlsx.InteropProbe.csproj` is the in-repo
example.)
