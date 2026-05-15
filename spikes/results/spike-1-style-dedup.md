# Spike 1 — Style-dedup feasibility

**Status:** Complete. Result reframes the design row rather than satisfying it.
**Date:** 2026-05-15
**NPOI version:** 2.7.3, .NET 9.0.314, linux-x64

## Question

What is the overhead of deduplicating cell styles via a per-workbook pool vs. not deduplicating? Original target (design §4): typical < 10%, worst case < 30%.

## Measurement

For each (cells × distinct-styles) scenario: build a workbook where every cell gets a style; in the **NoDedup** variant, a fresh `ICellStyle` is created per cell; in the **Dedup** variant, exactly *distinct* styles are pre-created and reused.

## Results

| Scenario              | NoDedup time | NoDedup pool | Dedup time | Dedup pool |
|-----------------------|--------------|--------------|-----------:|-----------:|
| 10k cells, 10 styles  | 0.44 s       | 10,001       | **0.03 s** | 11         |
| 100k cells, 10 styles | 10.6 s¹      | 60,001 (cap) | **0.28 s** | 11         |
| 500k cells, 10 styles | 10.2 s¹      | 60,001 (cap) | **0.83 s** | 11         |
| 100k cells, 100 styles| 10.5 s¹      | 60,001 (cap) | **0.14 s** | 101        |
| 100k cells, 1k styles | 10.5 s¹      | 60,001 (cap) | **0.23 s** | 1,001      |

¹ NoDedup hits NPOI's hard ~60,000-style cap before completing the workbook. The bench bails to record the partial result.

## Conclusion

**Dedup is not an "overhead" to budget — it is the only viable path.** Without dedup, any workbook with > ~60k styled cells hits NPOI's style-pool cap mid-write and the save fails. The original design target (< 10% / < 30% overhead vs raw NPOI) compared the wrong baselines: there is no "raw NPOI styling that goes higher than 60k cells" to be 10% slower than.

The right framing for the perf target is:

- **Throughput target:** styled writes scale roughly linearly with cell count, at ~600k styled-cells/s for a small palette and ~430k styled-cells/s for a 1k palette.
- **Capacity target:** style pool size equals the count of distinct fluent `CellStyle` values, not the count of styled cells. Confirmed: 10 distinct fluent styles → 11 pool entries (10 + the default).

## Design impact

| Row | Was | Now |
|-----|-----|-----|
| `design.md §4` "Style dedup overhead — typical / worst" | `< 10% / < 30% vs raw NPOI` | **Replaced** with capacity + throughput targets — see revised §4. |

The original two rows were measuring a phantom — "raw NPOI without dedup" cannot complete the workload, so percentage comparison is meaningless. Revising to honest capacity + throughput targets per spike-failure rule I21.

## Action items

- [x] Revise `design.md §4` perf-targets table.
- [x] Confirm style dedup remains v1.0 (already a Yes in roadmap — no change).
