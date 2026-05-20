# Benchmarks

BenchmarkDotNet harness exercising the v1.0 performance claims from
`docs/design.md §5`. Sized to complete in ~45 seconds on a modern
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

### CI cache (not committed; CI-hardware reference)

The `.github/workflows/bench.yml` workflow caches its run output keyed
by `hash(src/, benchmarks/*.cs, Directory.Packages.props)`. On a main
push, the cache for that key is populated with the run's results.
Subsequent PRs that touch a tracked path produce a new cache key, and
the workflow falls back via `restore-keys` to the closest previous
baseline.

The PR runs the comparator (`compare-bench.py`) against the cached
baseline and fails the build if any benchmark regresses more than
**15%**. The 15% threshold includes 5% headroom over the design DoD's
10% requirement for measurement noise on short-run CI benchmarks
(`ColdCreateAndSave` in particular has high variance because OPC
packaging dominates and varies with disk state).

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

| Benchmark | What it exercises | Design §5 target |
|---|---|---|
| `WriteBenchmarks.ColdCreateAndSave` | OPC packaging + workbook bootstrap | < 50 ms |
| `WriteBenchmarks.Write5kRows` | In-memory mixed-type write path | (proxy for 30k < 3s) |
| `WriteBenchmarks.StyledWrite_SmallPalette` | `CellStylePool` dedup (decision #4) | > 500k styled cells/s |
| `WriteBenchmarks.StreamingWrite_50kRows` | SXSSF streaming path | (proxy for 1M < 30s) |
| `ReadBenchmarks.OpenAndReadColumnSum` | Open + read | (proxy for 100k < 4s) |

The 5k / 50k / 1k variants are CI-sized scale-downs of the design-§5
targets, picked so a full run completes in under a minute on a typical
runner. Full-scale validation runs against the original §5 numbers are
done manually before each v1.0+ release (not in CI).
