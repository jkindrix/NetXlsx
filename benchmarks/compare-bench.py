#!/usr/bin/env python3
"""Compare BenchmarkDotNet brief-json results against a checked-in baseline.

Usage:
    python3 compare-bench.py <results-dir> <baseline-dir> [--threshold PCT]

Exit code 0 if no benchmark regresses beyond the threshold; 1 otherwise.

`results-dir` is BDN's `BenchmarkDotNet.Artifacts/results/`; we glob for
`*-report-brief.json` files inside it. `baseline-dir` is the same shape
(usually `benchmarks/baseline/`).

The comparison is per-benchmark, on the `Statistics.Mean` field
(nanoseconds). A "regression" is `(current - baseline) / baseline > threshold`.
Improvements (current < baseline) are reported but never fail the build.
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys
DEFAULT_THRESHOLD_PCT = 15.0  # design DoD says 10%; we add 5% headroom for
                              # measurement noise on short-run CI benchmarks
                              # (ColdCreateAndSave especially is high-variance)


def load_benchmarks(results_dir: str) -> dict[str, dict]:
    """Returns {FullName: {Mean, BytesAllocatedPerOperation}}."""
    benchmarks: dict[str, dict] = {}
    pattern = os.path.join(results_dir, "*-report-brief.json")
    for path in glob.glob(pattern):
        with open(path) as f:
            doc = json.load(f)
        for b in doc.get("Benchmarks", []):
            full = b["FullName"]
            mean_ns = b["Statistics"]["Mean"]
            allocated = b.get("Memory", {}).get("BytesAllocatedPerOperation", 0)
            benchmarks[full] = {"Mean": mean_ns, "Allocated": allocated}
    return benchmarks


def fmt_ns(ns: float) -> str:
    if ns < 1_000:
        return f"{ns:.0f} ns"
    if ns < 1_000_000:
        return f"{ns / 1_000:.2f} µs"
    if ns < 1_000_000_000:
        return f"{ns / 1_000_000:.2f} ms"
    return f"{ns / 1_000_000_000:.2f} s"


def fmt_bytes(b: int) -> str:
    if b < 1024:
        return f"{b} B"
    if b < 1024 ** 2:
        return f"{b / 1024:.1f} KiB"
    if b < 1024 ** 3:
        return f"{b / 1024 ** 2:.2f} MiB"
    return f"{b / 1024 ** 3:.2f} GiB"


def main(argv: list[str]) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("results_dir")
    p.add_argument("baseline_dir")
    p.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD_PCT,
                   help=f"Percent regression that fails the build (default: {DEFAULT_THRESHOLD_PCT}%%)")
    args = p.parse_args(argv)

    current = load_benchmarks(args.results_dir)
    baseline = load_benchmarks(args.baseline_dir)

    if not current:
        print(f"ERROR: no benchmark results found in {args.results_dir}", file=sys.stderr)
        return 2
    if not baseline:
        print(f"WARN: no baseline found in {args.baseline_dir} — recording current as new baseline", file=sys.stderr)
        return 0  # don't fail the first run; the workflow uploads the new baseline as an artifact

    print(f"Comparing {len(current)} current benchmark(s) against {len(baseline)} baseline(s)")
    print(f"Regression threshold: {args.threshold:.1f}%")
    print()

    regressions = 0
    new = 0
    removed = 0
    improvements = 0
    unchanged = 0

    all_keys = sorted(set(current.keys()) | set(baseline.keys()))
    print(f"{'Benchmark':<70} {'Baseline':>14} {'Current':>14} {'Δ%':>8}  Status")
    print("-" * 116)

    for k in all_keys:
        short = k.replace("NetXlsx.Benchmarks.", "")
        if k not in baseline:
            print(f"{short:<70} {'—':>14} {fmt_ns(current[k]['Mean']):>14} {'—':>8}  NEW")
            new += 1
            continue
        if k not in current:
            print(f"{short:<70} {fmt_ns(baseline[k]['Mean']):>14} {'—':>14} {'—':>8}  REMOVED")
            removed += 1
            continue
        b = baseline[k]["Mean"]
        c = current[k]["Mean"]
        pct = (c - b) / b * 100
        if pct > args.threshold:
            status = "REGRESSION"
            regressions += 1
        elif pct < -args.threshold:
            status = "improvement"
            improvements += 1
        else:
            status = "ok"
            unchanged += 1
        print(f"{short:<70} {fmt_ns(b):>14} {fmt_ns(c):>14} {pct:+7.2f}%  {status}")

    print()
    print(f"Summary: {regressions} regression(s), {improvements} improvement(s), "
          f"{unchanged} unchanged, {new} new, {removed} removed.")

    if regressions > 0:
        print(f"\nFAIL: {regressions} benchmark(s) regressed more than {args.threshold:.1f}%.", file=sys.stderr)
        print("To accept the regression as the new baseline, re-run benchmarks locally and "
              "commit the updated JSON files to benchmarks/baseline/.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
