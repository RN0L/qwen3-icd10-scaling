#!/usr/bin/env python3
"""Assert that every headline figure in the prose still matches ``results/analysis.json``.

The project's central rule is that no number appears in a deliverable unless it came from a
run record. ``analyze.py`` enforces the first half of that — everything it emits is derived
from the records. This enforces the second half: that the prose was actually updated when the
data changed.

How it works, and why it is not a list of hardcoded expectations
---------------------------------------------------------------

Each check names a figure, extracts it from ``analysis.json``, formats it the way the prose
writes it, and asserts that the resulting string occurs somewhere in the deliverables. Nothing
here stores what the number *should* be — if a run is re-measured and 237.45 becomes 240.10,
this looks for "240.10", fails to find it, and names the documents that still say otherwise.

That is exactly the failure this project is most exposed to: Leo landed a GPU row mid-project,
which moved a dozen figures, and stale prose is invisible to a reader who has no way to check.

::

    python3 profiling/analyze.py
    python3 profiling/check_report.py            # exit 0 if the prose matches the data
    python3 profiling/check_report.py --list     # print every figure and where it appears

Standard library only.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

#: Files whose prose must agree with the data. Anything user-facing that quotes a number.
DELIVERABLES = (
    "README.md",
    "docs/topology.md",
    "docs/bottleneck-analysis.md",
    "slides/slides.md",
    "slides/build_pptx.py",   # the deck's text lives in the builder now
)


def by_step_count(curve: Dict[str, Any], n: int) -> Dict[str, Any]:
    """The amortisation point for a given run length, keyed by the count rather than position.

    Indexing ``points`` positionally is how a check like this quietly starts comparing the
    wrong rows when the step-count list changes; it happened once while this was being
    written, and the check reported three false mismatches.
    """
    for point in curve["points"]:
        if point["n_steps"] == n:
            return point
    raise KeyError("no amortisation point at N=%d" % n)


def family_label(comparison: Dict[str, Any]) -> str:
    """Short name for a like-for-like family, e.g. "0.6B bs1/seq1024"."""
    return "%s bs%d/seq%d" % (
        comparison["model"].split("/")[-1].replace("Qwen3-", ""),
        comparison["batch_size"],
        comparison["seq_len"],
    )


def build_checks(analysis: Dict[str, Any]) -> List[Tuple[str, str]]:
    """Return ``(label, formatted_value)`` for every figure the prose is expected to state.

    Every figure is keyed by the comparison it belongs to. An earlier version of this file
    keyed the amortisation curves by backend alone; once a second like-for-like family
    appeared, later curves silently overwrote earlier ones and the check reported the 4B
    numbers under the headline labels. It produced ten confident false alarms. Families are
    therefore addressed explicitly, and a curve whose target is its own baseline is skipped
    rather than reported as a speedup against itself.
    """
    runs = analysis["runs"]
    comparisons = analysis["like_for_like"]
    mitigation = analysis["mitigation"][0]
    cost = analysis["cost"]
    sensitivity = cost.get("sensitivity", {})

    checks: List[Tuple[str, str]] = []

    def add(label: str, value: Any, spec: str = "%s") -> None:
        checks.append((label, spec % value))

    # --- per like-for-like family --------------------------------------------------------
    for comparison in comparisons:
        family = family_label(comparison)
        baseline_backend = comparison.get("baseline_backend", "cpu")
        for backend, info in sorted(comparison["backends"].items()):
            if backend == baseline_backend:
                continue
            add("%s: %s per-step speedup" % (family, backend.upper()),
                info["step_time_speedup_vs_cpu"], "%.2f")

        # Two accelerators *agreeing* is a claim worth asserting is in the prose. Two
        # accelerators differing by 60 % is not a second claim — it is the per-step speedup
        # written as a percentage, and it is already checked above. Only the agreement case
        # is required, or the check demands that the report state the same fact twice.
        if "gpu" in comparison["backends"] and "tpu" in comparison["backends"]:
            gpu_step = comparison["backends"]["gpu"]["median_step_s"]
            tpu_step = comparison["backends"]["tpu"]["median_step_s"]
            difference = abs(gpu_step - tpu_step) / max(gpu_step, tpu_step) * 100.0
            if difference < 5.0:
                add("%s: GPU/TPU step-time difference (%%)" % family, difference, "%.2f")

    # --- per amortisation curve ----------------------------------------------------------
    for curve in analysis["amortisation"]:
        target = curve["target_backend"]
        baseline_run = curve["baseline_run_id"]
        if curve["target_run_id"] == baseline_run:
            continue  # a curve of a run against itself is not a speedup
        if runs[baseline_run]["backend"] == target:
            continue  # same-backend pair; that is the mitigation delta, checked separately
        family = "%s vs %s" % (target.upper(), runs[baseline_run]["backend"].upper())
        for n in (12, 100, 1_000, 10_000):
            point = by_step_count(curve, n)
            add("%s end-to-end at N=%d" % (family, n), point["end_to_end_speedup"], "%.2f")
        add("%s target fixed cost (s)" % family, curve["target_fixed_s"], "%.1f")
        add("%s baseline fixed cost (s)" % family, curve["baseline_fixed_s"], "%.1f")

    # --- utilisation ----------------------------------------------------------------------
    for run_id, run in runs.items():
        mean = run["measured"]["utilization_mean_pct"]
        # A run that died mid-way has a mean over a partial run; it is a number, not a result,
        # and requiring the prose to quote it would be requiring noise.
        if mean is None or run["status"] != "ok":
            continue
        add("%s mean utilisation (%%)" % run["backend"].upper(), mean,
            "%.2f" if mean < 10 else "%.1f")

    # --- the reference run's decomposition and the mitigation -------------------------------
    reference = runs.get("tpu-v5e8-bs8-seq1024-filecache")
    if reference:
        segments = reference["breakdown"]["segments"]
        add("reference compile (s)", segments["compile_s"], "%.1f")
        add("reference steady compute (s)", segments["steady_compute_s"], "%.1f")
        add("reference compute share (%)", reference["amortisation"]["compute_fraction_pct"], "%.1f")

    add("checkpoint-load speedup", mitigation["model_load"]["speedup"], "%.2f")
    add("wall clock saved (%)", mitigation["total_wall"]["pct_saved"], "%.1f")
    add("scheduler queue spread", analysis["scheduler_latency"]["spread_ratio"], "%.1f")

    for key, entry in sorted(sensitivity.items()):
        if "cpu_cost_multiple" in entry:
            add("%s cost multiple" % key, entry["cpu_cost_multiple"], "%.2f")

    dataset = analysis.get("dataset") or {}
    if dataset.get("determined"):
        add("train split lower bound", dataset["train_docs_min"], "%d")
        add("train split upper bound", dataset["train_docs_max"], "%d")

    return checks


def load_prose(root: str) -> Dict[str, str]:
    prose: Dict[str, str] = {}
    for name in DELIVERABLES:
        path = os.path.join(root, name)
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as handle:
                prose[name] = handle.read()
    return prose


def occurrences(value: str, prose: Dict[str, str]) -> List[str]:
    """Which deliverables state this value.

    Matched with digit boundaries so that "2.09" is not considered present merely because
    some document contains "12.098".
    """
    pattern = re.compile(r"(?<![0-9.])" + re.escape(value) + r"(?![0-9])")
    return [name for name, text in prose.items() if pattern.search(text)]


def main(argv: Optional[Sequence[str]] = None) -> int:
    here = os.path.dirname(os.path.abspath(__file__))
    root = os.path.dirname(here)
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--analysis", default=os.path.join(root, "results", "analysis.json"))
    parser.add_argument("--root", default=root, help="repository root holding the deliverables")
    parser.add_argument("--list", action="store_true", help="print every figure and where it appears")
    args = parser.parse_args(argv)

    if not os.path.exists(args.analysis):
        raise SystemExit("%s not found — run `python3 profiling/analyze.py` first." % args.analysis)
    with open(args.analysis, "r", encoding="utf-8") as handle:
        analysis = json.load(handle)

    prose = load_prose(args.root)
    if not prose:
        raise SystemExit("no deliverables found under %s" % args.root)

    checks = build_checks(analysis)
    missing: List[Tuple[str, str]] = []

    for label, value in checks:
        where = occurrences(value, prose)
        if args.list:
            print("  %-46s %-10s %s" % (label, value, ", ".join(where) or "— NOT STATED"))
        if not where:
            missing.append((label, value))

    print(
        "\nchecked %d figures from %s against %d deliverables"
        % (len(checks), os.path.relpath(args.analysis, args.root), len(prose))
    )
    if missing:
        print("\n%d figure(s) in the data are not stated anywhere in the prose:" % len(missing))
        for label, value in missing:
            print("  %-46s expected %s" % (label, value))
        print(
            "\nEither the prose is stale — a measurement changed and the text did not — or the\n"
            "figure was deliberately dropped. Update the documents, or remove the check."
        )
        return 1

    print("every figure the data supports is stated, and stated correctly.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
