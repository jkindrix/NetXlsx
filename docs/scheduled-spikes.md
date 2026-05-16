# Scheduled spikes

Spikes that re-run on a cadence, not once-and-done. The pre-impl spikes
under `spikes/results/` answered their questions in 2026-05; the
questions below either (a) depend on upstream state that changes
independently or (b) are perf claims that should be re-verified as the
codebase grows.

The schedule is the load-bearing part. The decision rule in design.md
**I21** forbids silent target erosion — a missed schedule is a recorded
decision, not a drift.

---

## Spike 4-Q — NPOI AOT/trim posture re-check

**Question.** Does the *current* NPOI release still produce a runtime
failure under `PublishAot=true` and `PublishTrimmed=true`? (Spike 4
measured a hard `No` against NPOI 2.7.3 on 2026-05-15 — design decision
I2 / roadmap matrix marks AOT and Trim as `No†` with the retirement
condition: "promote when NPOI removes its `System.Xml.Serialization` /
`System.Reflection.Emit` dependencies, likely an NPOI 3.x change.")

**Method.** Bump the pinned NPOI version in a throwaway branch to the
latest stable release, re-run `spikes/NetXlsx.AotSpike/` end-to-end
in both `Mode=Aot` and `Mode=TrimOnly` configurations. Capture the
warning count, runtime behavior, and any new diagnostic IDs.

**Cadence.** Quarterly. Next due: **2026-08-16**. Subsequent: 2026-11-16,
2027-02-16, …

**Outcome record.** Append a dated row to the table below per run.

| Run date   | NPOI version | AOT runtime | Trim runtime | Notes |
|------------|--------------|-------------|--------------|-------|
| 2026-05-15 | 2.7.3        | Fails (`POIXMLException` at `XSSFWorkbook.OnWorkbookCreate`) | Fails (`POIXMLException` at `CreateSheet`) | Original spike 4. Roadmap rows set to `No†`. |
| 2026-08-16 | (pending)    | (pending)   | (pending)    | (pending) |

**Promotion / demotion rules.**
- If both AOT and Trim succeed at runtime: roadmap matrix rows lift from
  `No†` to `Yes` for the version of NetXlsx that ships the bump.
  Update decision I2.
- If Trim succeeds but AOT fails: split the row (Trim → `Yes`, AOT →
  `No†` retained). Update I2.
- If the bump regresses anything previously working: revert; record the
  regression; the schedule continues unchanged.

**Owner.** Whoever's holding the v1.0 implementation milestone at the
time. Not delegable to a build script — this is a "did the runtime
behavior actually change" question, not a "did the build pass"
question.

---

## Conventions

- One file per scheduled spike, additive history (don't overwrite past
  results — every run is data).
- A spike that produces a meaningful outcome change triggers a
  design-doc revision per process rule **I21**, not a silent table
  edit.
- A missed schedule is recorded as a CHANGELOG entry explaining why,
  not silently rolled forward.
