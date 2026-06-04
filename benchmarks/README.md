# Benchmarks

BenchmarkDotNet harness exercising the v1.0 performance claims from
`docs/design.md §4`. Sized to complete in ~45 seconds on a modern
laptop so the regression gate is fast enough for every PR.

## Layout

```
benchmarks/
├── NetXlsx.Benchmarks/
│   ├── Benchmarks.cs                 # CI-friendly [Benchmark]s + CiConfig
│   ├── Spike1_StyleDedupFeasibility.cs   # one-shot pre-impl spike (not a regression test)
│   ├── Spike2_StreamingBackPressure.cs   # same
│   ├── Spike3_AsyncWrappingCost.cs       # same
│   └── Program.cs                    # entry point: spike-N args -> spikes, else BDN switcher
├── baseline/                         # dev-local reference (faster hw, not CI's numbers)
│   ├── NetXlsx.Benchmarks.WriteBenchmarks-report-brief.json
│   └── NetXlsx.Benchmarks.ReadBenchmarks-report-brief.json
├── ci-baseline/                      # committed CI-hardware baseline — the bench.yml gate's comparand
└── compare-bench.py                  # comparator: regress > N% → exit 1
```

## Two baselines, two purposes

This project has **two separate baselines** because CI runners and dev
laptops have different hardware. Mixing them produces nonsense
comparisons.

### `benchmarks/baseline/` (committed; dev-local reference)

Captured on a developer machine for the **"did my change make things
slower on my hardware?"** check. Useful before opening a PR to catch
order-of-magnitude regressions early. Not used by CI.

Refresh manually after intentionally accepting a perf change:

```bash
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0 -- --filter '*Benchmarks*' --exporters JSON
cp BenchmarkDotNet.Artifacts/results/*-report-brief.json benchmarks/baseline/
# Inspect the diff, then commit if accepted.
```

Compare manually against the committed baseline:

```bash
python3 benchmarks/compare-bench.py BenchmarkDotNet.Artifacts/results benchmarks/baseline
```

### `benchmarks/ci-baseline/` (committed; CI-hardware reference)

The `.github/workflows/bench.yml` gate compares every run — main push,
PR, and dispatch alike — against the briefs **committed** in
`benchmarks/ci-baseline/` and fails the build if any benchmark
regresses more than **40%** (sub-5 ms benchmarks: **150%** —
`compare-bench.py`). There is no rolling
cache and no `continue-on-error`: a regression that lands on main keeps
failing the gate until it is either fixed or the baseline is refreshed
deliberately. A missing committed baseline fails the run loud rather
than silently recording the current run as the new truth.

History (I-87): the pre-v2.0.1 gate compared against an Actions cache
that refreshed on every main push and soft-failed its compare step
there — a 14.5× bulk-write regression shipped to the v2.0.0 tag with a
green Benchmarks badge. The committed-baseline design makes refresh a
reviewable git event instead of a workflow side effect.

Refresh deliberately, after accepting an intentional perf change:

1. Dispatch the Benchmarks workflow (`gh workflow run bench.yml`) or
   take any recent run on the desired commit.
2. Download its results artifact:
   `gh run download <run-id> -n bench-results-<run-number>`.
3. Copy the `*-report-brief.json` files into `benchmarks/ci-baseline/`,
   inspect the diff, and commit.

The thresholds (design DoD says 10%) are calibrated from an
identical-code A/B on the hosted runners (runs 26983785312 vs
26983940069, same engine source): macro benchmarks swung up to +27.6%
and sub-millisecond benchmarks up to +133% between runs. Hence the flat
line is 40% — it fails decisively on the ≥2× regression class the gate
exists for — and benchmarks whose baseline mean is under 5 ms gate at
150% (they are smoke checks; `ColdCreateAndSave` and the micro reads
vary with disk and runner state). If a failure looks like pure noise on
a test-stack-only bump, re-run the workflow before treating it as real.

## Running locally

```bash
# Full BDN suite (slow, statistically robust)
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0

# CI-style run (fast, ~45 sec total)
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0 -- --filter '*Benchmarks*' --exporters JSON

# Pre-impl spike harnesses (one-shot, not part of regression gate)
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0 -- spike-1
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0 -- spike-2
dotnet run --project benchmarks/NetXlsx.Benchmarks --configuration Release --framework net10.0 -- spike-3
```

## What's measured

| Benchmark | What it exercises | Design §4 target |
|---|---|---|
| `WriteBenchmarks.ColdCreateAndSave` | OPC packaging + workbook bootstrap | < 50 ms |
| `WriteBenchmarks.Write5kRows` | In-memory mixed-type write path | (proxy for 30k < 3s) |
| `WriteBenchmarks.StyledWrite_SmallPalette` | `CellStylePool` dedup (decision #4) | > 500k styled cells/s |
| `WriteBenchmarks.StreamingWrite_50kRows` | SXSSF streaming path | (proxy for 1M < 30s) |
| `ReadBenchmarks.OpenAndReadColumnSum` | Open + read | (proxy for 100k < 4s) |

The 5k / 50k / 1k variants are CI-sized scale-downs of the design-§4
targets, picked so a full run completes in under a minute on a typical
runner. Full-scale validation runs against the original §4 numbers are
done manually before each v1.0+ release (not in CI).
