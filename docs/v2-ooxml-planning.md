# v2.0 OOXML implementation — planning notes

This doc captures the research and time estimates for one of the v2
paths `docs/long-term.md` lists: implement OOXML directly inside
NetXlsx and drop the NPOI engine. Complementary to:

- `docs/npoi-3x-migration.md` — the v1.x "stay on NPOI but bump major"
  path.
- `docs/long-term.md` — the parent doc that puts this in context as
  the **last** of four adjacent options, ordered by current EV.

**Important framing (revised 2026-05-20 per external critique):**
the from-scratch path is **not** the headline alternative to staying
on NPOI 2.7.3. Per `docs/long-term.md`'s EV-ordered options, the
current best alternatives are (in order):

1. Bind ClosedXML as the engine.
2. Fork NPOI 2.7.3 under Apache-2.0.
3. Accept OSMF terms after an honest legal read.
4. From-scratch OOXML implementation (this doc).

This doc covers option 4 in detail because if we ever do pursue it,
the work is by far the largest and benefits most from research being
done early. **Reading the spec and competitor architectures is durable
knowledge regardless of which option we ultimately pick** — the
ClosedXML-bind path also benefits from understanding what ClosedXML
gets right and wrong about OOXML.

**Status (2026-05-20):** research only. No commitment to start before
v1.0 ships. The case for actually executing this path strengthens
only if AOT cleanness becomes a hard requirement for our users
*and* neither NPOI 3.x nor ClosedXML moves on AOT — see the AOT
contingency section below.

## The specification

ECMA-376 (5th edition, ISO/IEC 29500-equivalent) is free to download
from [ecma-international.org](https://ecma-international.org/publications-and-standards/standards/ecma-376/).
Mirrored at [github.com/QtExcel/ecma-376-5th](https://github.com/QtExcel/ecma-376-5th)
in PDF form for easier reference.

| Part | Size | Approx pages | Scope |
|---|---:|---:|---|
| 1 — Fundamentals + Markup Language Reference | 35 MB | ~5,000 | SpreadsheetML + WordprocessingML + PresentationML + DrawingML in one volume. The big one. |
| 2 — Open Packaging Conventions (OPC) | 1.9 MB | ~130 | The ZIP + content-types + relationships layer. Small but load-bearing — everything else builds on this. |
| 3 — Markup Compatibility and Extensibility | 0.9 MB | ~50 | The `mc:Ignorable` / `mc:AlternateContent` extension scheme. Not pivotal for v1.0-equivalent surface, but required for round-tripping files Excel emits. |
| 4 — Transitional Migration Features | 10 MB | ~1,500 | Differences between Strict (ISO-29500) and Transitional (what Excel actually writes) conformance. Excel writes Transitional; we need to read both. |

**For our surface specifically:** approximately 1,500-2,000 pages of
Part 1 (the SpreadsheetML chapters) + all of Part 2 (OPC) + the
Transitional-vs-Strict differences from Part 4 that affect Excel-
authored files. Estimate ~2,000 pages of mandatory reading; another
1,000-1,500 of "reference, look up when needed."

## Competitor projects

Verified facts from the GitHub API on 2026-05-20:

| Project | License | First commit | Stars | Repo size | Open issues | What to learn |
|---|---|---|---:|---:|---:|---|
| [**ClosedXML**](https://github.com/ClosedXML/ClosedXML) | MIT | 2016 (10 yr) | 5,594 | 73 MB | 427 | Most mature pure-OOXML in .NET; primary reference. Read for architecture, stylesheet handling, perf optimizations. The 10 years of consumer-feedback baked into its API design is the value. |
| [**NPOI**](https://github.com/nissl-lab/npoi) | Apache-2.0 | 2013 (13 yr) | 6,173 | 162 MB | many | What we currently wrap. Read for what's hard for a port-from-Java — clues for what *not* to copy. |
| [**Open-XML-SDK**](https://github.com/dotnet/Open-XML-SDK) | MIT | 2014 (12 yr) | 4,516 | 88 MB | n/a | Microsoft's low-level OOXML SDK. Not a high-level library — raw schema binding. Use as a *correctness oracle* for our output. |
| [**EPPlus**](https://github.com/EPPlusSoftware/EPPlus) | Custom (paid ≥5.0) | 2019 branch | 2,020 | 48 MB | mostly paid | Older history goes back to ~2009 on a different repo. Pre-5.0 versions on NuGet are MIT and worth reading. Strong on commercial-feature edge cases. |
| [**MiniExcel**](https://github.com/mini-software/MiniExcel) | Apache-2.0 | 2021 (5 yr) | n/a (API quirk) | 47 MB | low | Smallest scope, fastest streaming. Read for *what you can skip* — they prove a focused subset still has real users. |
| [**xlsxwriter** (Python)](https://github.com/jmcnamara/XlsxWriter) | BSD-2 | 2013 (13 yr) | external | — | — | Different ecosystem, but the cleanest "write a single sheet, save" architecture in the field. Read for layering ideas. |

## Time estimate

NetXlsx's v1.0 contract is **narrower** than ClosedXML's full surface.
We expose: write + read + style + freeze/merge + named ranges +
formulas (no eval) + comments + hyperlinks + streaming. We *do not*
author pivot tables, charts, or conditional formatting from scratch —
we just preserve them as opaque OPC parts. That cuts approximately
30-40% off the implementation surface vs a "full ClosedXML clone."

| Phase | Solo, focused full-time | Solo, ~10 hr/wk (with day job) |
|---|---:|---:|
| Spec reading + design sketches | 4-8 weeks | 3-6 months |
| MVP write (1 sheet, text + number → save → opens in Excel) | 2-4 months | 6-9 months |
| Read + preservation (open arbitrary .xlsx, round-trip unmodeled parts) | 2-3 months | 4-6 months |
| Styling + advanced surface (matches current `IRange`/`IColumn`/`ICell`) | 3-4 months | 6-9 months |
| Streaming write (SXSSF analog) | 2-3 months | 4-6 months |
| Hardening + fuzz tests + real-world corpus | 3-6 months | 6-12 months |
| **Total to "ready as v2.0 default engine"** | **12-18 months** | **24-36 months** |

**Caveats on the estimate (revised 2026-05-20 per external critique):**

- **The hardening phase is probably underestimated by 2-3×.** Excel
  emits invalid OOXML routinely. LibreOffice emits *different*
  invalid OOXML. Google Sheets has its own quirks. The graveyard
  of half-finished OOXML libraries is large for a reason — the
  spec is 5,000 pages and reality is a superset of the spec.
  ClosedXML has 10 years and 427 open issues' worth of "how to
  not throw on this weird file" baked in; we'd be re-paying that
  tuition. Realistic hardening phase: 6-18 months full-time, not
  3-6.
- Multi-contributor speedup is sublinear. Two contributors splitting
  spec-reading + MVP write probably doesn't halve the timeline —
  realistic best case is ~30% reduction.
- The "as v2.0 default engine" bar means it passes the existing
  434+ test/TFM suite *and* the preservation fixture's four-part-
  type round-trip *and* the bench gate without > 15% regression.
  Skipping any of those bars cuts time but compromises the v1.0
  contract guarantee.
- **Solo-maintainer permanence.** A 50k-LOC OOXML engine is a
  permanent maintenance commitment: every CVE in the XML parsing
  chain, every new Excel-version quirk, every conditional-formatting
  edge case lands on the maintainer. NPOI exists because a team
  maintains it; ClosedXML exists because a team maintains it. A
  one-person fork-or-rebuild becomes a bus-factor cliff. (See the
  bus-factor honesty paragraph in `docs/long-term.md` — it's a
  delta on top of the wrapper's existing bus-factor, not a new
  category of risk, but a real delta.)
- **Opportunity cost.** In 18 months of v2 reimplementation, the
  v1.1 roadmap could land Tables, data validation, image
  embedding, rich text, themes, autofilter — all high-user-value
  features that go faster on the existing engine. The from-scratch
  path defers all of that.

## AOT contingency — when from-scratch becomes the right answer

The from-scratch case is dormant *unless* AOT cleanness becomes a
hard requirement for a meaningful chunk of our users. If that
happens, the EV calculation flips: from-scratch becomes the only
path that delivers AOT, since both NPOI and ClosedXML use
`System.Xml.Serialization` + reflection patterns that the AOT
compiler can't satisfy.

**Active signals we're watching** (folded into the existing quarterly
`Spike 5-Q`):

1. **NPOI 3.x's AOT posture.** If NPOI 3.x ships AOT-clean, the
   from-scratch case weakens further — bind NPOI 3.x and we get
   AOT for free.
2. **ClosedXML's AOT posture.** Same logic. As of 2026-05, neither
   engine is AOT-clean and neither has announced AOT work.
3. **Consumer-side demand signal.** Specific signals to weight
   heavily:
   - Issues filed on NetXlsx asking for AOT/trim support.
   - .NET ecosystem trends: Native AOT in API hosts (e.g.
     ASP.NET Core 8+ minimal APIs), Blazor WASM trimming
     pushing toward mandatory, .NET MAUI mobile (iOS App Store
     basically requires AOT).
   - Direct consumer requests from teams who tried NetXlsx and
     found the AOT block disqualifying.

**Trigger.** If two or more of these signals fire within the same
quarterly Spike 5-Q review *and* neither NPOI 3.x nor ClosedXML
has announced AOT work, the from-scratch path moves from "option 4"
to "option 1" in `docs/long-term.md`'s ordering. That's a real
re-baselining event recorded under the new "Roadmap re-baselining"
process — not silent drift.

Until then, the from-scratch option stays parked. Don't let the
research investment in this doc create pressure to execute the
path — the doc's job is to keep the option open and to make the
trigger condition explicit, not to make execution inevitable.

## Study sequence (pre-implementation)

Independent of when the implementation work itself starts, the
*reading* phase is cheap to begin now. Suggested sequence:

1. **Read first, code second.** Spend at least a month with Part 1
   (SpreadsheetML chapters) and Part 2 (OPC) end-to-end before
   writing any v2 code. Annotate. Identify which sections of Part 1
   we actually need for our v1.0 surface.
2. **Study ClosedXML's architecture** as the primary reference.
   It's MIT-licensed; we can read freely. Focus on:
   - How they layer OOXML serialization vs the public API.
   - Their stylesheet handling — the analog of our `CellStylePool`.
   - Performance optimizations (especially around large workbooks).
   - Pain points they hit (mine the GitHub issues — see below).
3. **Skim NPOI source** for the SAME surface, comparing approaches.
   Differences between ClosedXML's from-scratch and NPOI's
   port-from-Java reveal *opportunities* — places one chose
   differently from the other.
4. **Use Open-XML-SDK as a correctness oracle.** Build a test that
   takes our v2 output, parses with Open-XML-SDK, asserts the
   schema validates. Catches "we wrote subtly wrong XML" before
   Excel does.
5. **Mine GitHub issues** for "OOXML pain points":
   - ClosedXML's 427 open + thousands closed: search `data loss`,
     `corruption`, `preserve`, `OOM`, `large file`, `pivot`. These
     tell you where the hard parts of OOXML actually live (often
     not where the spec says they do).
   - EPPlus security advisories — real attacks against `.xlsx`
     parsers (XML expansion bombs, zip-slip, malformed relationships).
   - NPOI's issue tracker — what NPOI itself struggles with on
     the .NET side.
6. **Build a real-world corpus.** Microsoft has a public corpus of
   "interesting" .xlsx files from their compatibility testing.
   Synthesize one from publicly-available datasets (government
   open-data .xlsx files, Wikipedia exports, etc.) — a few hundred
   files spanning Excel-authored, LibreOffice-authored, and
   Google-Sheets-authored.

## Decision shape

Per `docs/long-term.md`'s revised R&D milestones, the decision to
*start* the from-scratch implementation isn't pre-committed. Two
gates have to pass:

1. **Spike-comparison gate** (`docs/long-term.md` R&D-1 step):
   the parallel from-scratch + bind-ClosedXML single-sheet-write
   spikes produce comparable measures. If the bind-ClosedXML
   spike covers our v1.0 facade with bounded effort and acceptable
   peer-dependency posture, this gate selects bind-ClosedXML and
   the from-scratch path stays parked. The from-scratch spike's
   output becomes a documented "we considered and rejected"
   artifact rather than the input to a continuation.
2. **AOT contingency gate** (this doc's AOT-contingency section):
   if AOT becomes binding (per the trigger conditions above) *and*
   neither NPOI 3.x nor ClosedXML has moved toward AOT support,
   the from-scratch case re-strengthens regardless of what the
   spike-comparison gate said.

The study sequence above can begin **today** regardless of which
gate eventually triggers execution — reading the OOXML spec and
the competitor architectures is durable knowledge for the
bind-ClosedXML path too (we'd need to understand what ClosedXML
gets right and wrong about OOXML to translate effectively).

## Status

- **Current** (2026-05-20): research-only. Not started. From-scratch
  is option 4 of 4 in `docs/long-term.md`'s EV-ordered list.
- **Next milestone:** post-v1.0-tag, run the R&D-1 parallel spikes
  (from-scratch + bind-ClosedXML at same scope). The output
  selects between continuing from-scratch (R&D-2/R&D-3),
  switching to bind-ClosedXML, or shelving both and re-evaluating
  NPOI 3.x.
