# Spike 4 — AOT / trim posture

**Status:** Complete. Result invalidates the optimistic design row.
**Date:** 2026-05-15
**Project:** `spikes/NetXlsx.AotSpike/`
**NPOI version:** 2.7.3
**Runtime:** .NET 9.0.314, linux-x64

## Question

Does NPOI 2.7.3 survive `PublishAot=true` and/or `PublishTrimmed=true`?

## Measurement

Minimal `XSSFWorkbook` create → set string/number/date cells → save → re-open → read back, published as a self-contained Linux x64 binary, then executed.

## Results

### Mode 1 — `PublishAot=true` (full AOT)

- **Compile:** succeeds.
- **Warnings:** 56 (38 × `IL3050` AOT analysis + 18 × `IL2xxx` trim analysis). Located in `System.Private.Xml`, `EnumsNET`, and `NPOI`.
- **Binary size:** 15.7 MB (single self-contained native binary; no JIT).
- **Runtime:** **FAILS** with `POIXMLException` inside `NPOI.POIXMLDocumentPart.CreateRelationship`, during `XSSFWorkbook.OnWorkbookCreate`. NPOI's OOXML relationship subsystem relies on reflection/dynamic code that AOT cannot satisfy.

### Mode 2 — `PublishTrimmed=true` (trim only, JIT preserved)

- **Compile:** succeeds.
- **Warnings:** 4 × `IL2104` "assembly produced trim warnings" (one each for `Enums.NET`, `NPOI.Core`, `NPOI.OOXML`, `NPOI.OpenXmlFormats`). Per-method warnings are aggregated.
- **Binary size:** 75 KB (small — most of NPOI was trimmed away, but incorrectly).
- **Runtime:** **FAILS** at `CreateSheet("AotSpike")` with the same `POIXMLException`. Trimming removed code paths NPOI's reflection needs at runtime.

## Conclusion

**Both AOT and trim are incompatible with NPOI 2.7.3 today.** The failure is in NPOI's OOXML serialization layer, which depends on `System.Xml.Serialization` (heavy runtime codegen) and on reflection over NPOI's own type graph. These are not fixable from outside NPOI.

## Design impact (per I21 spike-failure handling)

| Row | Was | Now |
|-----|-----|-----|
| Roadmap "Native AOT compatible" v1.0 | `TBD*` | **`No`** |
| Roadmap "Trim compatible" v1.0 | `TBD*` | **`No`** |
| Decision I2 "AOT/trim posture" | "TBD pending AOT spike; likely outcome: trim-with-warnings; AOT-incompatible" | "AOT-incompatible and trim-incompatible while NPOI uses runtime XML serialization. Promote when NPOI ships AOT-clean serialization, or descope NPOI in favor of an OOXML library that does." |
| `docs/design.md §10` ("Definition of extraordinary") | "Performance per the targets in §4" — silent on AOT | Add: "AOT/trim status is honestly stated in the roadmap; the facade does not pretend to AOT-readiness while its engine is AOT-hostile." |

This is a **structural** outcome per the spike-failure rule (§3.7 of methodology / I21): the optimistic design row was wrong; the roadmap matrix needs a real `No` in both cells. The library itself does not change — the facade layer is AOT-clean by construction; the runtime dependency makes the whole package AOT-hostile.

No descope of NPOI itself is warranted: AOT is not a stated v1.0 requirement, just an aspirational posture. Typical consumers run JIT.

## Action items

- [x] Update `docs/roadmap.md` matrix: AOT row → `No`; trim row → `No`. Footnote retained but reframed.
- [x] Update `docs/design.md` decision I2 to reflect measured outcome, not predicted.
- [x] Update CHANGELOG `[Unreleased]` to note the spike outcome.

## Retirement condition

This conclusion stands until NPOI removes its dependency on `System.Xml.Serialization` and `System.Reflection.Emit` paths. Track NPOI 3.x; revisit the AOT spike when a major bump lands.
