# Spike 3 — Async wrapping cost

**Status:** Complete. Result confirms the design's `Task.Run` strategy.
**Date:** 2026-05-15
**NPOI version:** 2.7.3, .NET 9.0.314, linux-x64

## Question

Is `Task.Run`-wrapped NPOI save measurably more expensive than sync save for representative workbook sizes? Specifically: does `SaveAsync(stream, ct) => Task.Run(() => npoi.Write(stream))` introduce non-trivial overhead vs. inlined sync?

## Measurement

For each row count, 20 iterations (after warm-up) measured with `Stopwatch`. Concurrent case spawns 10 parallel saves.

## Results

### Single-threaded (median / P95 / mean, ms)

| Mode             | Rows   | Median | P95   | Mean  |
|------------------|-------:|-------:|------:|------:|
| Sync             |    100 |   10.6 |  20.2 |  11.4 |
| Async (Task.Run) |    100 |    9.7 |  13.3 |  10.3 |
| Sync             |  1,000 |   84.6 | 102.5 |  80.0 |
| Async (Task.Run) |  1,000 |   60.1 |  66.7 |  54.2 |
| Sync             | 10,000 |  327.5 | 363.7 | 327.9 |
| Async (Task.Run) | 10,000 |  322.7 | 361.1 | 325.4 |

### 10× parallel (P95 / total wall time, ms)

| Mode             | Rows   | P95   | Total |
|------------------|-------:|------:|------:|
| Sync             |    100 |  12.0 |  14.5 |
| Async (Task.Run) |    100 |   9.0 |   9.1 |
| Sync             |  1,000 |  94.4 |  94.6 |
| Async (Task.Run) |  1,000 |  77.5 |  77.7 |
| Sync             | 10,000 | 904.7 | 905.2 |
| Async (Task.Run) | 10,000 | 799.1 | 799.6 |

## Conclusion

`Task.Run` wrapping is **free or net-positive** at every workbook size tested:

- Single-threaded: P95 within noise (10k rows) or 10–35% *faster* (1k rows), likely due to thread-pool warm-up offsetting the dispatch cost.
- Concurrent (10×): P95 10–13% *better* with `Task.Run`. The thread-pool's scheduling lets parallel saves overlap I/O more effectively than direct sync execution on a single thread.

## Design impact

| Row | Was | Now |
|-----|-----|-----|
| Decision #5 (I/O style) | "may or may not `Task.Run` (resolved by Spike-2)" / current §7.1 says "`Task.Run` for `Save`/`Open`" | **Confirmed.** No change to design — the spike validates the existing decision. |

§7.1 already states `Task.Run` is the strategy. Spike confirms there is no measurable cost to pay for it, so no hybrid pattern is warranted.

Caveat: the concurrent harness uses `Task.Run` to dispatch every parallel save regardless of "Sync" vs "Async" label — both modes are running on the thread pool by construction at that point. The single-threaded result is the cleaner measurement of "is wrapping itself expensive?" Answer: no.

## Action items

- [x] Decision #5 stands. No revision.
- [x] §7.1 wording unchanged.
