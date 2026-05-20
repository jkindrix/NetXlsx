# NetXlsx — long-term direction

This is the doc for tracked R&D work that is real but post-v1.0.
Items here are deliberate aspirations with framing tight enough that
they could become a milestone with a date and an owner — they're not
free-form wishlist entries.

## OOXML implementation from scratch — v2.0 R&D track

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

**R&D track milestones (sequence; not yet dated):**

1. **v1.0 ships** on NPOI 2.7.3. Get consumer adoption. Measure the
   surface area that's actually exercised. (Status: pre-v1.0, three
   ship-blockers remaining.)
2. **v1.x maintenance** continues on the 2.7.3 pin. Issues filed
   against the wrapper layer get fixed; NPOI quirks documented as
   they're encountered. (Status: continuous from v1.0.)
3. **R&D-1: Native OOXML write spike.** Implement the smallest
   end-to-end surface (write a single sheet with text + number cells,
   save) by directly serializing the OOXML package + workbook.xml +
   sharedStrings.xml + sheet1.xml. Match the observable behavior of
   the current `Workbook.Create().AddSheet(...).Save(...)` path
   under the same `IWorkbook` facade. Lands as
   `Workbook.CreateNative()` returning `IWorkbook`. Measures:
   complexity, time-to-MVP, remaining-surface-area, AOT cleanliness,
   golden-file byte-comparable output. Spike, not production work.
4. **R&D-2: Native OOXML read spike.** Same scope inversion — implement
   the smallest open/read path under the same `IWorkbook` facade.
   `Workbook.OpenNative(...)` returning `IWorkbook`. Same measures.
5. **R&D-3: Coverage matrix.** Iterate R&D-1 and R&D-2 surface
   coverage until the facade's full v1.0 contract is satisfied. Run
   the existing 433-test suite (golden + unit) against both the
   NPOI engine and the native engine through the same `IWorkbook`
   interface; any divergence is a bug in the native engine.
6. **Decision point.** When the native engine passes the existing
   test suite, decide: ship as v2.0 (deprecate NPOI engine), ship
   alongside as opt-in (`Workbook.CreateNative()` stays explicit),
   or shelve as completed R&D and re-evaluate NPOI 3.x.

The decision at step 6 is the load-bearing one and we don't pre-commit
to it. The R&D delivers data; the decision uses it.

**Investment shape.** R&D-1 alone is bounded — a single-sheet
text+number write is < 2k LOC of OOXML serialization. The R&D track
becomes expensive at R&D-3 (full coverage); that's where the real
multi-quarter commitment lives. Each step gates the next.

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
