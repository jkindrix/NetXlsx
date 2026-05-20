# NetXlsx — long-term direction

This is the parking-lot doc for ideas that are real but not v1.0 work.
Items here are deliberate aspirations, not commitments. They live here
so the design discipline doesn't lose them.

## OOXML implementation from scratch (post-v1.0)

NetXlsx is a wrapper over NPOI today. That was the right call for
v1.0 — it let us focus on the ergonomics layer (fluent API, style-pool
dedup, source-generator typed mapping, escape hatch) without
re-implementing OOXML. But it leaves us tied to NPOI's release cadence,
license posture (see decision **I23** for the OSMF situation), API
quirks (workbook-name-uniqueness, `date1904` hardcoded, etc.), and the
ongoing AOT/trim block (decision **I2**).

**The aspiration.** Implement OOXML directly inside NetXlsx and drop
the NPOI dependency entirely.

**Why this is the right direction (eventually):**

- Every other major .NET spreadsheet library — ClosedXML, EPPlus,
  MiniExcel — implements OOXML directly. NPOI is the .NET community's
  outlier "port from Java POI" approach.
- The OSMF situation makes the NPOI dependency a licensing risk that
  will keep coming up at quarterly re-evaluation (`docs/scheduled-spikes.md`).
- AOT/trim compatibility is gated entirely on NPOI's
  `System.Xml.Serialization` + `System.Reflection.Emit` usage. A
  ground-up OOXML implementation could be AOT-clean from day one
  (design choice).
- The wrapper layer (which is most of NetXlsx's actual ergonomic value)
  is *not* coupled to NPOI internals — it's coupled to the OOXML
  semantics we expose. Swapping the engine under that layer is a real
  amount of work but not architecturally invasive.

**Why this is not v1.0 work:**

- OOXML is a 5,000+ page ECMA-376 specification (Parts 1–4). Full
  conformance is a multi-quarter project at minimum.
- The current NPOI 2.7.3 pin is workable; v1.0 has real consumer value
  *now* without doing the from-scratch work first.
- Scope discipline: each consumer who picks up v1.0 informs what
  OOXML surface area is actually used in practice. Implementing the
  long tail before that signal arrives is a recipe for over-investing.

**Plausible sequencing (not a commitment):**

1. **v1.0 ships** on NPOI 2.7.3. Get consumer adoption. Measure the
   surface area that's actually exercised.
2. **v1.x maintenance** continues on the 2.7.3 pin. Issues filed
   against the wrapper layer get fixed; NPOI quirks documented as
   they're encountered.
3. **Spike: OOXML write path MVP.** Pick the smallest surface (write
   a single sheet with text + number cells, save) and implement the
   OOXML package + workbook.xml + sharedStrings.xml + sheet1.xml
   serialization directly. Match the observable behavior of the
   current `Workbook.Create().AddSheet(...).Save(...)` path. Measure
   complexity, time-to-MVP, and remaining-surface-area honestly.
4. **Decision point.** Based on the spike: continue full
   reimplementation as a parallel `NetXlsx.Native` engine, or
   reverse the call and accept the NPOI dependency long-term.

The decision at step 4 is the load-bearing one and we don't pre-commit
to it. The spike makes it informed.

**Adjacent options that aren't full reimplementation:**

- **Fork NPOI at 2.7.3** and maintain it under Apache-2.0. Keeps the
  engine; trades the consumer of NPOI for the maintainer of an NPOI
  fork. Cost is real (NPOI is ~200K LOC) but bounded — only the
  surface NetXlsx actually exercises needs to keep working.
- **Bind a different engine.** ClosedXML's underlying OOXML
  implementation is the most NPOI-comparable option. License is
  MIT, scope is comparable. The translation layer becomes a different
  kind of work, not less work.
- **Accept the OSMF terms** and document them loud in the README.
  Reverses I23. Requires a real legal read on the transitive
  obligation question.

None of these are pre-decided. The quarterly spike (`Spike 5-Q` in
`scheduled-spikes.md`) keeps the option open.

## Other long-term items

(Add new sections here as they come up. Format: `## Title` + 1-2
paragraphs of context + a clear "this is not v1.0 work" statement.)
