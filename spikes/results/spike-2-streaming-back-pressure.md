# Spike 2 — Streaming back-pressure

**Status:** Complete. In-memory target missed; streaming target on track.
**Date:** 2026-05-15
**NPOI version:** 2.7.3, .NET 9.0.314, linux-x64

## Question

At what row count does in-memory `XSSFWorkbook` exceed the §4 working-set target? Does `SXSSFWorkbook` (streaming) hold flat memory across 500k+ rows?

## Measurement

For each scenario: build a workbook of *rows × cols* (20 cols, mixed cell types — strings and numbers), save, capture deltas vs scenario baseline.

- **ΔWS** = `Process.WorkingSet64` peak minus baseline (post-GC, pre-scenario).
- **ΔGC** = `GC.GetTotalMemory(false)` peak minus baseline.

## Results

### In-memory (`XSSFWorkbook`)

| Rows    | Time   | ΔWS     | ΔGC     | File   |
|--------:|-------:|--------:|--------:|-------:|
| 10,000  |  1.5 s | 184 MB  | 122 MB  | 0.9 MB |
| 50,000  |  2.5 s | 614 MB  | 631 MB  | 4.4 MB |
| 100,000 |  5.1 s | 956 MB  | 1.27 GB | 8.8 MB |
| 250,000 | 12.0 s | 2.6 GB  | 3.1 GB  | 22 MB  |

### Streaming (`SXSSFWorkbook`, 100-row window)

| Rows    | Time   | ΔWS    | ΔGC     | File   |
|--------:|-------:|-------:|--------:|-------:|
| 10,000  |  0.5 s |   0 MB |  46 MB  | 0.7 MB |
| 50,000  |  1.4 s |  11 MB |  48 MB  | 3.5 MB |
| 100,000 |  2.5 s |   0 MB |  36 MB  | 7.1 MB |
| 250,000 |  6.2 s |   0 MB |  37 MB  | 18 MB  |
| 500,000 | 12.8 s |  76 MB |  69 MB  | 35 MB  |

## Conclusion

In-memory write **scales linearly at ~10–12 GB managed-memory per million cells**. Streaming write **stays flat at ~50–70 MB managed-memory regardless of row count**, exactly the back-pressure shape SXSSF is built to produce.

Against design §4 targets:

| Target                                                  | Measured              | Status |
|---------------------------------------------------------|-----------------------|--------|
| In-memory 100k rows × 20 cols < 3s, < 500 MB            | 5.1 s, 956 MB ΔWS    | **Miss** |
| Streaming 1M rows × 20 cols < 30s, < 200 MB             | extrap. ~26 s, ~140 MB | On track (linear extrapolation from 500k) |
| Cold create empty workbook + save < 50 ms               | not measured here     | Spike 3 measured ~10 ms |
| Open + read 100k × 20 sheet < 4 s                       | not in this spike     | Read-side measurement deferred |

## Design impact (per I21)

The in-memory 100k-row target is unachievable on NPOI 2.7.3. Three options:

- **Revise the target** (chosen). Lower the in-memory promise to a row count where it actually holds; recommend streaming above that.
- Implement a workaround — none available; the memory cost is in NPOI's `XSSFWorkbook` object graph itself.
- Descope — would mean removing in-memory write, which is unacceptable.

Revised target row in design §4:
- In-memory 30k rows × 20 cols: < 3 s, < 500 MB. Hits the budget with margin.
- Above 30k rows: callers should use the streaming entry point (`CreateStreaming`).

This threshold matches the cookbook's existing "TabularExport" (10k rows) and "StreamingMillionRows" recipes — the recommendation falls cleanly between them.

## Action items

- [x] Revise `design.md §4` row "Write 100k rows × 20 cols (in-memory)" to "Write 30k rows × 20 cols (in-memory) < 3 s, < 500 MB".
- [x] Add a note to the streaming-recommendation: streaming is preferred above ~30k rows.
- [x] Roadmap row "Streaming write (`IStreamingWorkbook`)" remains v1.0 (no change — it was already there).
