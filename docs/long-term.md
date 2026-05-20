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
3. **R&D-1: parallel single-sheet-write spikes.** Two bounded spikes
   at identical scope, run in parallel for a real apples-to-apples
   comparison:
   - **R&D-1a (from-scratch native):** implement the smallest
     end-to-end surface (write a single sheet with text + number
     cells, save) by directly serializing the OOXML package +
     workbook.xml + sharedStrings.xml + sheet1.xml. Lands as a
     prototype `Workbook.CreateNative()` behind the same
     `IWorkbook` facade. Bounded at ~2k LOC of OOXML serialization.
   - **R&D-1b (bind-ClosedXML):** implement the same surface by
     binding ClosedXML's `XLWorkbook` behind the same `IWorkbook`
     facade. Lands as a prototype `Workbook.CreateOnClosedXML()`.
     Bounded at probably 2-3 weeks of translation work.

   Both spikes produce a golden-file output for the same input.
   Measures captured for each: complexity (LOC + cyclomatic), time
   to MVP, remaining-surface-area-to-v1.0-contract, AOT cleanliness
   (run under `PublishAot=true` and check for runtime trim warnings),
   byte-comparable golden output, peer-dependency posture.

   This is the load-bearing experiment. The original framing of
   "R&D-1 → R&D-2 → R&D-3 from-scratch only" was a sunk-cost ladder
   that compared from-scratch against itself; this parallel design
   produces an actual decision input.

4. **Decision point — gate.** Compare R&D-1a vs R&D-1b on the
   measures above. Branch:
   - If R&D-1b (bind-ClosedXML) covers the v1.0 facade with
     bounded effort and acceptable peer-dependency posture:
     **adopt the bind-ClosedXML path**, ship as v2.0, deprecate
     the NPOI engine. The from-scratch spike (R&D-1a) becomes
     a documented "we considered and rejected" artifact.
   - If R&D-1b reveals load-bearing facade methods that don't
     translate cleanly to ClosedXML's API, *and* R&D-1a shows
     the from-scratch effort is tractable at the v1.0 scope:
     proceed to R&D-2/R&D-3 from-scratch (steps 5-6 below).
   - If both spikes reveal the v1.0 facade is too far ahead of
     either engine: stay on NPOI 2.7.3 indefinitely, treat the
     v2.0 R&D as completed-and-shelved, re-evaluate NPOI 3.x
     when it ships.

5. **R&D-2 (from-scratch path only): Native OOXML read spike.** Same
   scope inversion — implement the smallest open/read path.
   Reached only if step 4's gate selects the from-scratch branch.

6. **R&D-3 (from-scratch path only): Coverage matrix.** Iterate
   R&D-1a + R&D-2 surface coverage until the facade's full v1.0
   contract is satisfied. Run the existing 434-test suite against
   both the NPOI engine and the native engine through the same
   `IWorkbook` interface; any divergence is a bug in the native
   engine. Reached only if R&D-2 also passes.

**Investment shape.** R&D-1a + R&D-1b together are ~5-6 weeks of
bounded work. Step 4's gate is where most projects fail to be
ruthless — be ruthless. R&D-1b passing the gate is the most likely
outcome by EV and it ends the R&D track cleanly; only continue if
the spike output actually argues for from-scratch on its measures,
not because the original plan said R&D-2 came next.

**Adjacent options — ordered by current expected value, not by
authorial preference.** Per the external critique on 2026-05-20, the
original ordering led with from-scratch and treated alternatives as
afterthoughts. That under-weighted the EV math. Honest ordering:

1. **Bind ClosedXML as the engine.** MIT-licensed, .NET-native, mature
   (10 years, 5.6k stars, no Java-port baggage), no OSMF risk. The
   translation layer is bounded work (probably weeks, not years)
   because most of our facade methods map to a single ClosedXML
   call. Cost: we reintroduce the "load-bearing single-vendor
   upstream" pattern with a different vendor — ClosedXML has had
   governance gaps historically and we'd be at the mercy of its
   maintainers' priorities. Does **not** unlock AOT (ClosedXML uses
   the same `System.Xml.Serialization` + reflection patterns that
   block NPOI under AOT). This is probably the highest-EV option
   today *unless* AOT becomes a hard requirement — see below.
2. **Fork NPOI at 2.7.3** and maintain under Apache-2.0. Keeps the
   engine; trades being a consumer of NPOI for being the
   maintainer of an NPOI fork. NPOI is ~200K LOC but ~half is HSSF
   (.xls) + HWPF (Word) which we don't use — realistic maintenance
   surface for our facade is closer to 30-50k LOC. Bounded, but
   only the right move if the community fork option (Spike 5-Q
   trigger #3) doesn't materialize first.
3. **Accept the OSMF terms** with an honest legal read on whether the
   transitive obligation actually applies to wrapper-library
   consumers. Reverses I23. Cheapest path *if* the legal read
   concludes consumers of NetXlsx aren't transitively obligated by
   NPOI's OSMF EULA. We treated this as last-resort in the original
   doc without doing the legal read — that was lazy. It might
   survive scrutiny.
4. **Full from-scratch OOXML implementation.** Last, not first. The
   case strengthens significantly *only if* AOT cleanness becomes
   a hard requirement for a meaningful chunk of users (Native AOT
   API hosts, Blazor WASM trimming, .NET MAUI) and neither NPOI 3.x
   nor ClosedXML moves toward AOT compatibility in a reasonable
   horizon. In that case, from-scratch would make us the only
   AOT-clean option in the field — a real differentiator. Without
   that contingency, the EV is dominated by the 12-36-month
   timeline against opportunity cost (v1.1 features that go faster
   on the existing engine: Tables, data validation, images, rich
   text, themes, autofilter).

The quarterly spike (`Spike 5-Q` in `scheduled-spikes.md`) keeps the
options open — it specifically watches for the AOT-demand signal,
the NPOI 3.x trigger, and any community-fork emergence.

**Bus-factor honesty.** A from-scratch engine is ~50k LOC the
maintainer owns forever — CVEs in the XML parsing chain, every
Excel-version quirk, every conditional-formatting edge case. That's
a real risk. But the wrapper itself is already solo-maintained; if
the maintainer disappears tomorrow, the wrapper goes unmaintained
regardless of which engine sits underneath. From-scratch makes the
bus-factor *worse* (more code, more domain knowledge required to
take over) but doesn't introduce it where it didn't exist.

**Anchoring honesty.** The original framing of this doc led with
from-scratch as the headline because the maintainer had stated
intent to "go all-in" on it. Writing a plan to support a stated
intent isn't planning — it's commitment laundering. This revision
removes that anchor; the from-scratch option still exists but it's
where the EV math actually puts it.

## Roadmap re-baselining (post-v1.0)

The roadmap matrix in `docs/roadmap.md` has three states that aren't
binary "Yes" or "Never": **v1.1** / **v2.0** / **v3.0** (scheduled
intent) and **Deferred†** (intent contingent on an upstream unblock).
Without a structured re-look, those rows drift — features get
scheduled for "v1.1" and then quietly stay there release after
release; "Deferred†" rows lose their unblock-trigger context.

**Cadence.** After v1.0 ships, the matrix is re-baselined twice a
year (alongside the existing quarterly spike re-checks):

1. **Q1 review** (Jan/Feb): every non-Yes/non-Never row gets one of
   three verdicts:
   - **Promote** — move to the next-earlier release column (e.g.,
     v1.1 → v1.0 patch line if there's a real consumer ask).
   - **Demote** — move later (v1.1 → v2.0) or to Never if the row's
     premise has eroded. Demotions require a one-line rationale in
     CHANGELOG.
   - **Hold** — keep the current column but add a date-stamped
     "still expected" note. A row in Hold for 4 consecutive reviews
     (2 years) auto-demotes to v3.0 or Never.
2. **Q3 review** (Jul/Aug): same process; lighter cadence between
   reviews. Adjacent to the Spike 4-Q + Spike 5-Q quarterly re-checks
   so review sessions amortize.

**Method.** A markdown checklist generated from `docs/roadmap.md`'s
matrix is added to a tracking issue. Each row gets a one-line verdict.
The matrix gets edited in-place; the checklist gets archived on the
tracking issue.

**Deferred† rows have stronger rules.** Each Deferred† row names the
unblock trigger explicitly (the AOT/trim rows name "NPOI 3.x removes
its problematic deps OR our native OOXML engine lands"). If the
trigger event happens, Promote/Demote in the next review. If the
trigger event doesn't happen for 8 consecutive reviews (4 years),
demote to Never — the deferral has become indefinite, which isn't a
state we keep.

**This is not v1.0 work.** It's the v1.x+ maintenance discipline.
Recorded here so the matrix doesn't accumulate drift between major
releases. First scheduled review after v1.0 tag: the **next** Jan/Feb
or Jul/Aug, whichever is closer.

## Other long-term items

(Add new sections here as they come up. Format: `## Title` + 1-2
paragraphs of context + a clear "this is not v1.0 work" statement.)
