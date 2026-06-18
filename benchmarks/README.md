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

## Comparative results (R-29)

Everything above measures NetXlsx **in isolation** against the absolute
design-§4 targets. The comparative suite (`Comparative.cs`, ledger R-29)
adds the missing dimension: NetXlsx against the two 2026 reference
points named in the ledger —

- **SpreadCheetah** — the closest analog (streaming / source-gen / AOT /
  MIT). Write-only and streaming-only, so it is compared **only on the
  streaming-write path** (its README's 100k×10 ≈ 33 ms reference point).
- **ClosedXML** — the breadth comparator (full DOM read + write).
  Compared on the **buffered-write** and **read** paths.

SpreadCheetah is never pitted on read or buffered-DOM features it does
not have; ClosedXML is never pitted on a streaming path it does not have.

### Out of the CI gate, by construction

The comparative classes live in the **`NetXlsx.Comparative`** namespace,
not `NetXlsx.Benchmarks`. The bench.yml regression gate selects with
`--filter "*Benchmarks*"` and BenchmarkDotNet's filter matches the full
benchmark id (namespace included), so the comparative classes are never
run — and never compared — by the gate. This is deliberate (decision S3 /
I-87): a competitor's run-to-run variance must never be able to red our
own absolute-number gate. They run only on an explicit filter:

```bash
# Real SDK on PATH first (BDN spawns a child build that must satisfy global.json):
PATH="$HOME/.dotnet:$PATH" DOTNET_ROOT="$HOME/.dotnet" \
  dotnet run --project benchmarks/NetXlsx.Benchmarks -c Release \
  --framework net10.0 -- --filter '*Comparative*'
```

The `ComparativeConfig` already wires the GitHub-markdown and brief-JSON
exporters, so no `--exporters` flag is needed (passing one duplicates the
exporter and BDN warns).

### Measured (2026-06-18)

A single local run on an otherwise-idle workstation — **12th Gen Intel
Core i9-12900, .NET SDK 10.0.102, .NET 10.0.2 runtime, Linux** — under
the suite's `ShortRun` job (3 warmup × 3 iterations). The 99.9 %
confidence intervals are wide purely as an N=3 artifact; run-to-run
`StdDev` is tight (≤ ~2 % on the buffered/read pairs, ~0.5 % on NetXlsx
streaming). The gaps below are order-of-magnitude and dwarf that noise,
so `ShortRun` is enough to substantiate the comparison — these are not
precision microbenchmarks, and the absolute milliseconds will differ on
your hardware. Re-run before reading anything into a sub-2× difference.

**Streaming write — 100,000 rows × 10 mixed-scalar cols (= 1M cells):**

| Engine                | Mean       | Allocated  | vs NetXlsx          |
|-----------------------|-----------:|-----------:|---------------------|
| NetXlsx (streaming)   | 1,309 ms   | 1,037 MB   | baseline            |
| **SpreadCheetah**     | **50.3 ms**| **4.1 MB** | **~26× faster, ~255× less allocated** |

**Buffered write — 50,000 rows × 10 mixed-scalar cols (= 500k cells):**

| Engine                | Mean       | Allocated  | vs NetXlsx          |
|-----------------------|-----------:|-----------:|---------------------|
| NetXlsx (buffered)    | 1,668 ms   | 620 MB     | baseline            |
| **ClosedXML**         | **704 ms** | **481 MB** | **~2.4× faster, ~0.8× allocated** |

**Read — open the 50,000-row × 10-col file, sum one numeric column:**

| Engine                | Mean       | Allocated  | vs NetXlsx          |
|-----------------------|-----------:|-----------:|---------------------|
| NetXlsx               | 373 ms     | 160 MB     | baseline            |
| **ClosedXML**         | **276 ms** | **150 MB** | **~1.35× faster, ~0.9× allocated** |

### Reading these honestly

NetXlsx is **slower than both reference points on every path measured.**
This is the comparative reality R-29 existed to surface, and it is stated
plainly rather than buried:

- **Streaming write is the largest and most material gap.** SpreadCheetah
  is a purpose-built forward-only writer with a source-generated cell path
  and almost no intermediate allocation; NetXlsx's streaming path runs
  through the Open XML SDK's `OpenXmlWriter` and currently churns ~1 GB of
  transient objects to emit 1M cells (≈255× SpreadCheetah's allocation).
  The ~1 GB figure is total managed *allocation throughput* over the run,
  not retained working set — it does not by itself contradict §4's
  "< 200 MB ΔWS" streaming target (a working-set, not allocation, claim) —
  but the allocation profile is a genuine efficiency signal and is flagged
  for operator triage (a candidate roadmap perf item, **not** in R-29's
  measurement-only scope).
- **Against ClosedXML — which shares the very same Open XML SDK engine —**
  NetXlsx is ~2.4× slower buffered and ~1.35× slower on read. NetXlsx's
  thinner object model is **not** currently buying a speed advantage over
  the breadth library; ClosedXML's mature write path is faster here.
- **The absolute §4 targets still hold** when these scaled workloads are
  extrapolated (see `docs/design.md §4`). What this data retires is any
  implicit "thin ⇒ fast" assumption. NetXlsx's differentiation is
  **architectural** — thinness over the SDK, the raw-SDK escape hatch,
  compile-time-checked typed mapping, the dedup style pool, and
  preservation fidelity — **not raw throughput.** The top-level README's
  positioning is already worded that way (it claims streaming write as a
  *capability*, never as *faster than* a competitor); this data confirms
  that wording was the right call.
