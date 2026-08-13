# The derived contract — what is in `results/analysis.json`

[`metrics-contract.md`](metrics-contract.md) governs what a *run* writes. This file governs
what the *analysis* writes, because there are now two contracts in this repository and only
one of them was documented.

The chain the report depends on is:

```
results/<run_id>.json      measured, one per run, metrics-contract.md
        │
        │  profiling/analyze.py      ← the only place derived numbers are computed
        ▼
results/analysis.json      derived, one file, this document
        │
        ├── profiling/make_dashboard.py  → results/systems_dashboard.png
        ├── slides/build_pptx.py          → the deck
        ├── README.md, docs/*.md          (prose, written by hand from these fields)
        └── profiling/check_report.py     → asserts the prose still matches
```

Nothing downstream of `analysis.json` recomputes anything. If a number in the report cannot be
traced to a field below, it is a bug — and `check_report.py` exists to catch the specific case
where a field changed and the prose did not.

---

## Top level

| Key | What it holds |
|---|---|
| `generated_by` | the script that wrote the file |
| `generated_from` | results directory, record count, and the run IDs consumed |
| `validation` | what was checked, what warned, and which fields are not measurements |
| `runs` | every run, keyed by `run_id`, with its measured and derived values |
| `like_for_like` | cross-backend comparisons where the workload configuration is identical |
| `amortisation` | one curve per accelerator: `time(N) = fixed_s + N · median_step_s` |
| `sweeps` | `batch_size` and `seq_len` families, including the OOM boundary |
| `mitigation` | before/after pairs differing only in `file_cache_capacity` |
| `scheduler_latency` | spread of `submit_to_running_s` across accelerator runs |
| `cost` | per-run cost at the pinned list prices, with sensitivity |
| `dataset` | training split size inferred from epoch-wraparound notes (`docs/dataset_stats.json` is now committed and is authoritative) |
| `coverage` | what the measurement set does and does not contain |

---

## `validation` — the part that decides whether anything else may be used

```json
{
  "all_records_satisfy_contract": true,
  "warnings": {"<run_id>": ["..."]},
  "tainted_fields": {"<run_id>": ["steps[].loss"]},
  "plotted_fields": ["steady", "memory", "phases", "utilization"],
  "checks_applied": ["..."]
}
```

- **`tainted_fields`** is the field-scoped result of reading each record's `notes`. A record
  saying its per-step probe may have blocked on a stale buffer taints `steady`; a record
  saying no per-step loss could be captured taints `steps[].loss` and nothing else.
- **A taint on a field in `plotted_fields` is a hard error**, not a warning: `analyze.py`
  refuses to emit output. This exists because one earlier run reported 11 million tokens/s
  with `status: "ok"` and was caught only by a human reading `notes`.
- All nine records currently carry the loss taint and none is thereby untrustworthy for step
  times. A whole-record verdict would have lost that distinction.

---

## `runs.<run_id>` — the refined decomposition

Beyond echoing `config`, `hardware` and the contract's own `phases`, each run carries:

```json
"breakdown": {
  "segments": {
    "scheduler_queue_s": 116.0,     "image_pull_s": null,
    "process_init_s": 86.1,         "checkpoint_load_s": 16.1,
    "compile_s": 138.5,             "straggler_s": 35.3,
    "steady_compute_s": 108.6,      "checkpoint_write_s": 2.1
  },
  "segment_pct": { "...": "each as a percentage of total_wall_s" },
  "unattributed_s": 5.7e-14,
  "sums_to_total": true,
  "detail": { "second_step_ratio": 39.3, "second_step_is_recompile": true, "...": "..." }
}
```

**Why these differ from `phases`.** The contract defines `steady_state_s` as `sum(steps[1:])`,
which still contains the second step — a second XLA compile on every accelerator record here —
and the long steps at epoch and checkpoint boundaries. Charting that as compute would credit
the accelerator with time it spent compiling. The split uses an identity exact for any `m`:

```
sum(S) = n·m + (S[0] − m) + (S[1] − m) + Σ_{i≥2}(S[i] − m)
```

with `m = steady.median_step_s`. `unattributed_s` is the reassembly residual and is ≤ 1.1e-13 s
on every record; `sums_to_total` asserts it. The last term is **signed** — steps faster than
the median subtract — so nothing is invented and nothing is lost.

`segments.compile_s` merges both compilations when `detail.second_step_is_recompile` is true
(ratio ≥ 10× the median); otherwise the second step is charged to `straggler_s`. The ratio that
produced the judgement is recorded so a reader who disagrees can move the time back.

```json
"amortisation": {
  "fixed_s": 394.2, "marginal_s_per_step": 1.75123,
  "compute_fraction_pct": 21.6, "overhead_fraction_pct": 78.4,
  "fixed_excl_scheduler_s": 278.2
}
```

`fixed_s` is **not fitted**: it is the measured wall clock minus every recorded step charged at
the steady rate. `fixed_excl_scheduler_s` exists because the CPU plane has no scheduler at all,
so including queue time charges the accelerators for something the baseline never pays.

`other_s_composition` recovers the four named costs `telemetry.py` writes into `notes`
(`jax_init`, `lora_wrap`, `data_prep`, `trainer_build`); `other_s_composition_residual_s` is
what they do not account for — interpreter startup and imports, which nothing times explicitly.
It is reported rather than absorbed, so the segment is not read as fully explained.

---

## Comparisons

**`like_for_like`** only groups runs whose `model`, `batch_size` and `seq_len` all match. There
are three such families: the 0.6B trio at batch 1, and 4B at batch 8 and 16 where the GPU and
the TPU meet. The CPU is the baseline where it has a record; where it does not — it cannot run
4B at all — the baseline falls back to the slowest backend present, and `baseline_backend` and
`baseline_is_cpu` say which happened. Read `baseline_is_cpu` before quoting a speedup: the 4B
figures are GPU-against-TPU, not against the CPU. Each backend entry carries
`step_time_speedup_vs_cpu`, `speedup_per_chip`, and `strong_scaling_efficiency: null` with the
reason attached — a cross-architecture ratio divided by a chip count is not a parallel
efficiency, and no single-chip TPU baseline was measured.

**`amortisation`** is one curve per (family, accelerator) pair — six of them, not two — each
with `points` at N = 12 … 100 000,
`asymptotic_speedup`, and `breakeven_n_steps`. A **negative** break-even means the
accelerator's fixed cost is already below the baseline's — it wins at every run length, and
`breakeven_note` says so rather than leaving a negative number to be read as a threshold.

**`sweeps`** is keyed per backend — each entry carries its own `backend` field, and `held` holds
only the dimensions actually held constant. It defaulted to the TPU until the 4B GPU sweep was
measured, at which point four records contributed nothing and a v5e-specific finding was being
quoted as a general one. Records with `status: "oom"` are kept: the failure boundary is the
point of the sweep, so `max_ok` / `min_oom` bound it. `scaling` gives the exponent between consecutive points
(`t ~ x^α`), which is how "step time grows as `seq^1.30` while peak memory stays flat" is
stated without eyeballing a chart.

**`mitigation`** pairs runs on the `config` block rather than on the `-filecache` suffix in the
`run_id`, so a rename cannot silently drop the comparison. `steady_step.pct_change` is the
control: a storage fix that also moved compute time would mean the two runs differed in
something else.

---

## `cost` — three rates, one of which is real

`prices` pins each rate with a source URL, a retrieval date, and
`applies_to_measured_hardware`. **Only the TPU rate is true of hardware that was actually
billed.** The CPU and GPU rates are proxies for on-premises machines with no list price, and
the GPU carries a second substitution besides: Google Cloud does not sell a GH200 at all.

Every per-run entry is reported twice, `usd_per_1000_steps` and `usd_per_1000_steps_excl_queue`,
because the inclusive figure contains whatever scheduler queue that run happened to draw and
queue time varies by more than an order of magnitude. Configurations are compared on the
exclusive column. `sensitivity` states, for each conclusion, how far the price would have to
move before it reversed.

---

## `coverage` — generated, so the report cannot overstate itself

`gaps` is built from the records rather than written by hand: the "no GPU record" gap
disappeared by itself when the GPU row landed. Each gap carries `what`, `short` (used in the
dashboard), `consequence`, `explained_in` and `kind`. The README's Known-gaps table and the
dashboard's bottom strip are both rendered from this list, which is why they cannot drift apart
or claim coverage the records do not have.
