# v2.0 OOXML implementation — planning notes

This doc captures the research and time estimates for the v2.0 path
that `docs/long-term.md` calls out: implement OOXML directly inside
NetXlsx and drop the NPOI engine. Complementary to
`docs/npoi-3x-migration.md` (the v1.x path that keeps NPOI as the
engine but adopts NPOI 3.x once trigger conditions fire).

**Status (2026-05-20):** research only. No commitment to start before
v1.0 ships.

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

**Caveats on the estimate:**

- The hardening phase is the one most rebuilds underestimate. It's
  where you encounter Excel-authored files that violate the spec
  (Excel itself does this regularly) but that consumers expect to
  round-trip without throwing.
- Multi-contributor speedup is sublinear. Two contributors splitting
  spec-reading + MVP write probably doesn't halve the timeline —
  realistic best case is ~30% reduction.
- The "as v2.0 default engine" bar means it passes the existing
  433+ test/TFM suite *and* the preservation fixture's four-part-
  type round-trip *and* the bench gate without > 15% regression.
  Skipping any of those bars cuts time but compromises the v1.0
  contract guarantee.

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

Per `docs/long-term.md`'s R&D milestones, the decision to *start*
the v2.0 implementation work isn't pre-committed. The trigger is:

- v1.0 is in consumer hands long enough that the actually-used
  surface is known (probably 6-12 months post-tag).
- The NPOI 3.x situation has resolved one way or the other (either
  3.x ships with terms we accept, or it doesn't ship at all by
  some reasonable horizon).
- Project owner commits the time-budget (full-time vs day-job-side
  changes the timeline by 2-3x — see table above).

The study sequence above can begin **today** regardless of
implementation start. Reading the spec and competitor architectures
is durable knowledge whether or not we ever write a line of v2 code.

## Status

- **Current** (2026-05-20): research-only. Not started.
- **Next milestone:** post-v1.0-tag, project owner decides whether
  to begin the study sequence in parallel with v1.x maintenance
  or wait until v1.x is in stable maintenance mode.
