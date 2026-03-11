#!/usr/bin/env python3
"""
Benchmark script for solver_logic.py.
Runs optimization with different category_diversity_weight values and records results.

Usage:
    python benchmarks/bench_solver.py [--label "version label"]
"""
import sys
import os
import time
import argparse
import tempfile
from datetime import datetime

# Add webapp to path so we can import solver_logic
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "webapp"))
from solver_logic import run_optimization

INPUT_FILE = os.path.join(os.path.dirname(__file__), "..", "archive", "filled_template_2025_with_categories.xlsx")
RESULTS_FILE = os.path.join(os.path.dirname(__file__), "results.md")
WEIGHTS = [0, 5, 10, 15]


def run_bench(weight, input_path):
    """Run a single benchmark with the given category weight. Returns (time_s, stats_dict, success, msg)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        output_path = f.name
    try:
        t0 = time.time()
        success, msg, stats = run_optimization(input_path, output_path, category_diversity_weight=weight)
        elapsed = time.time() - t0
        return elapsed, stats, success, msg
    finally:
        if os.path.exists(output_path):
            os.unlink(output_path)


def format_table(results):
    """Format results as a markdown table."""
    header = "| Weight | Time (s) | Objective | Pref Rate | Vetoes | Neutrals | Cat Diversity |"
    sep = "|--------|----------|-----------|-----------|--------|----------|---------------|"
    rows = [header, sep]
    for r in results:
        cat_div = r.get("cat_div", "-")
        rows.append(
            f"| {r['weight']:>6} | {r['time']:>8.1f} | {r['objective']:>9} | {r['pref_rate']:>9} "
            f"| {r['vetoes']:>6} | {r['neutrals']:>8} | {cat_div} |"
        )
    return "\n".join(rows)


def main():
    parser = argparse.ArgumentParser(description="Benchmark solver performance")
    parser.add_argument("--label", default="unlabeled", help="Version label for this benchmark run")
    args = parser.parse_args()

    if not os.path.exists(INPUT_FILE):
        print(f"ERROR: Input file not found: {INPUT_FILE}")
        sys.exit(1)

    print(f"=== Solver Benchmark ({args.label}) ===")
    print(f"Input: {INPUT_FILE}")
    print(f"Weights: {WEIGHTS}")
    print()

    results = []
    for w in WEIGHTS:
        print(f"--- Running weight={w} ---")
        elapsed, stats, success, msg = run_bench(w, INPUT_FILE)
        if not success:
            print(f"  FAILED: {msg}")
            results.append({
                "weight": w, "time": elapsed, "objective": "FAIL",
                "pref_rate": "-", "vetoes": "-", "neutrals": "-", "cat_div": "-"
            })
            continue

        cat_div = ""
        if stats.get("category_diversity_distribution"):
            sorted_keys = sorted(stats["category_diversity_distribution"].keys(),
                                 key=lambda k: int(k.split("/")[0]), reverse=True)
            cat_div = ", ".join(f"{k}:{stats['category_diversity_distribution'][k]}" for k in sorted_keys)

        results.append({
            "weight": w,
            "time": elapsed,
            "objective": stats.get("objective_value", "N/A"),
            "pref_rate": stats.get("pref_rate", "N/A"),
            "vetoes": stats.get("veto_count", "N/A"),
            "neutrals": stats.get("neutral_count", "N/A"),
            "cat_div": cat_div or "-",
        })
        print(f"  Time: {elapsed:.1f}s | Obj: {stats.get('objective_value')} | Pref: {stats.get('pref_rate')}")
        print()

    # Print table
    table = format_table(results)
    print("\n" + table + "\n")

    # Append to results.md
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(RESULTS_FILE, "a") as f:
        f.write(f"\n## {args.label} ({timestamp})\n\n")
        f.write(table + "\n")

    print(f"Results appended to {RESULTS_FILE}")


if __name__ == "__main__":
    main()
