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

## Spike 5-Q — NPOI OSMF posture re-check

**Question.** Has the NPOI OSMF (Open Source Maintenance Fee, introduced
in 2.8.0) situation changed in a way that would let us bump past 2.7.3?
Specifically:

- Have the OSMF terms been clarified to exclude transitive (wrapper)
  dependencies?
- Has the OSI weighed in on whether the OSMF EULA preserves OSI-approved
  status?
- Has the NPOI community forked a clean-Apache-2.0 maintained line?
- Has the long-term own-OOXML work in NetXlsx progressed to where we
  could drop NPOI entirely?

**Method.** Read the current OSMF site (`opensourcemaintenancefee.org/consumers/`),
the NPOI GitHub discussions, the OSI license-discuss archives, and any
SPDX or FSF positions. Summarize. Decide whether decision **I23** holds,
needs amendment, or needs reversal.

**Cadence.** Quarterly. Next due: **2026-08-16** (aligned with
Spike 4-Q to amortize the review session). Subsequent: 2026-11-16, …

**Outcome record.** Append a dated row to the table below per run.

| Run date   | NPOI latest | OSMF status            | I23 verdict        | Notes |
|------------|-------------|------------------------|--------------------|-------|
| 2026-05-20 | 2.8.0       | EULA on binary; transitive obligations unclear | **Hold pin at 2.7.3** | Initial assessment. |
| 2026-08-16 | (pending)   | (pending)              | (pending)          | (pending) |

**Promotion rules.**
- If OSMF terms exclude transitive dependencies *and* the community
  accepts that interpretation: lift the pin to latest 2.x. Update I23.
- If a clean Apache-2.0 maintained fork emerges with comparable scope:
  switch to it. Document the migration.
- If the own-OOXML work has reached MVP coverage of the v1.0 surface:
  start the deprecation conversation for the NPOI dependency.
- Otherwise: pin stands.

**Owner.** Project owner (license/business call; not a build-script
question).

---

## Conventions

- One file per scheduled spike, additive history (don't overwrite past
  results — every run is data).
- A spike that produces a meaningful outcome change triggers a
  design-doc revision per process rule **I21**, not a silent table
  edit.
- A missed schedule is recorded as a CHANGELOG entry explaining why,
  not silently rolled forward.
