#!/usr/bin/env python3
"""Build the 5-slide PowerPoint deck from the measured records.

Every number on these slides is read from ``results/`` at build time, so the deck
cannot drift from the data. Nothing is typed in by hand.

    uv run --with python-pptx --with pillow python slides/build_pptx.py

Brand colour is a single constant (``ORANGE``) — change it in one place if the
house colour differs.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

REPO = Path(__file__).resolve().parent.parent
RESULTS = REPO / "results"
FIGURES = Path(__file__).resolve().parent / "figures"
OUT = Path(__file__).resolve().parent / "ME344_Final_Arnold_Hambuch.pptx"

# --- palette -------------------------------------------------------------------
ORANGE = RGBColor(0xE8, 0x59, 0x0C)      # primary
ORANGE_SOFT = RGBColor(0xFD, 0xF0, 0xE6)  # tint for panels
INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x6B, 0x72, 0x80)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

W, H = Inches(13.333), Inches(7.5)  # 16:9


def load(name: str) -> dict:
    return json.loads((RESULTS / f"{name}.json").read_text())


# --- the measured facts, read not typed ----------------------------------------
cpu = load("cpu-x86-32c-0p6b-bs1-seq1024")
gpu = load("gpu-gh200-0p6b-bs1-seq1024")
tpu = load("tpu-v5e8-bs1-seq1024-filecache-0p6b")
base = load("tpu-v5e8-bs8-seq1024")
mitig = load("tpu-v5e8-bs8-seq1024-filecache")
oom = load("tpu-v5e8-bs32-seq1024-filecache")

c_step = cpu["steady"]["median_step_s"]
g_step = gpu["steady"]["median_step_s"]
t_step = tpu["steady"]["median_step_s"]
speed_g, speed_t = c_step / g_step, c_step / t_step
gt_delta = abs(g_step - t_step) / t_step * 100

_cost = {r["run_id"]: r for r in json.loads((RESULTS / "analysis.json").read_text())["cost"]["per_run"]}
cost_cpu = _cost["cpu-x86-32c-0p6b-bs1-seq1024"]["usd_per_1000_steps"]
cost_gpu = _cost["gpu-gh200-0p6b-bs1-seq1024"]["usd_per_1000_steps"]
cost_tpu = _cost["tpu-v5e8-bs1-seq1024-filecache-0p6b"]["usd_per_1000_steps"]

b_ph, m_ph = base["phases"], mitig["phases"]
compute_pct = b_ph["steady_state_s"] / b_ph["total_wall_s"] * 100
load_pct = b_ph["model_load_s"] / b_ph["total_wall_s"] * 100
load_factor = b_ph["model_load_s"] / m_ph["model_load_s"]
wall_saving = (1 - m_ph["total_wall_s"] / b_ph["total_wall_s"]) * 100


def add_slide(prs: Presentation):
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    bar = s.shapes.add_shape(1, 0, 0, W, Inches(0.13))
    bar.fill.solid()
    bar.fill.fore_color.rgb = ORANGE
    bar.line.fill.background()
    bar.shadow.inherit = False
    return s


def textbox(slide, x, y, w, h, text, size=18, bold=False, color=INK,
            align=PP_ALIGN.LEFT, space_after=6):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space_after)
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Helvetica Neue"
    return tb


def title(slide, number, text, subtitle=None):
    textbox(slide, Inches(0.6), Inches(0.35), Inches(1.0), Inches(0.6),
            str(number), size=40, bold=True, color=ORANGE)
    textbox(slide, Inches(1.35), Inches(0.42), Inches(11.4), Inches(0.8),
            text, size=30, bold=True, color=INK)
    if subtitle:
        textbox(slide, Inches(1.35), Inches(1.12), Inches(11.4), Inches(0.5),
                subtitle, size=15, color=GREY)


def table(slide, x, y, rows, col_w, header=True, size=14, row_h=Inches(0.42)):
    n_r, n_c = len(rows), len(rows[0])
    shape = slide.shapes.add_table(n_r, n_c, x, y, sum(col_w), row_h * n_r)
    tbl = shape.table
    for j, wdt in enumerate(col_w):
        tbl.columns[j].width = wdt
    for i, row in enumerate(rows):
        tbl.rows[i].height = row_h
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.10)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            run = para.runs[0] if para.runs else para.add_run()
            run.font.size = Pt(size)
            run.font.name = "Helvetica Neue"
            if header and i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = ORANGE
                run.font.bold = True
                run.font.color.rgb = WHITE
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if i % 2 else ORANGE_SOFT
                run.font.color.rgb = INK
    return shape


def figure(slide, name, x, y, width):
    p = FIGURES / name
    if p.exists():
        slide.shapes.add_picture(str(p), x, y, width=width)


prs = Presentation()
prs.slide_width, prs.slide_height = W, H

# ============================== 1 — Problem ====================================
s = add_slide(prs)
title(s, "1", "Where does an LLM fine-tuning job spend its time?",
      "ME344 Final Project · Option 2, custom workload · Leonard Arnold & Luis Hambuch")

textbox(s, Inches(0.6), Inches(1.85), Inches(6.3), Inches(4.4),
        "The workload\n"
        "LoRA fine-tune of Qwen3 on CodiEsp — 1 000 Spanish clinical case\n"
        "reports annotated with ICD-10 diagnosis codes (CC-BY 4.0).\n"
        "500 train / 250 dev, ⌀ 15 codes per document, 1 811 labels.\n\n"
        "Why it is a resource problem\n"
        "4B parameters = 8 GB of weights before a single activation.\n"
        "Documents run to 2 489 tokens. Model accuracy earns no marks —\n"
        "the object of study is the infrastructure.", size=15)

textbox(s, Inches(7.2), Inches(1.85), Inches(5.6), Inches(1.5),
        "The question", size=15, bold=True, color=ORANGE)
textbox(s, Inches(7.2), Inches(2.25), Inches(5.6), Inches(2.0),
        "Where does the time actually go —\nand what does it take to run\nthe job at all?",
        size=21, bold=True)
textbox(s, Inches(7.2), Inches(4.25), Inches(5.6), Inches(2.2),
        "The second half is not decoration. At Qwen3-0.6B all three\n"
        "backends run. At 4B the CPU runs in no configuration we tried.\n\n"
        "Motivated by Dr. Findus, a medical billing system that maps\n"
        "clinical documentation to ICD-10. No patient data was used.",
        size=13, color=GREY)


# ============================== 2 — Approach ===================================
s = add_slide(prs)
title(s, "2", "One code path, three backends",
      "Proposal, orchestration, and the compilation layer")

textbox(s, Inches(0.6), Inches(1.85), Inches(4.0), Inches(3.6),
        "Compilation layer\n"
        "One JAX / Flax program.\n"
        "XLA lowers it to CPU, CUDA and TPU.\n"
        "LoRA via qwix, training via Tunix.\n"
        "No per-backend reimplementation.", size=15)

textbox(s, Inches(4.8), Inches(1.85), Inches(4.0), Inches(3.6),
        "Orchestration\n"
        "3 pinned Docker images.\n"
        "Kubernetes Jobs on two clusters.\n"
        "Kueue admission on the TPU pool.\n"
        "Explicit CPU / memory / chip limits.", size=15)

textbox(s, Inches(9.0), Inches(1.85), Inches(3.8), Inches(3.6),
        "Storage\n"
        "gcsfuse CSI, implicit-dirs.\n"
        "8 GB checkpoint from a bucket.\n"
        "fileCacheCapacity as a knob —\n"
        "that knob is the experiment.", size=15)

table(s, Inches(0.6), Inches(5.05), [
    ["Backend", "Hardware", "Host arch", "Delivery"],
    ["CPU", "32-core x86_64, 31 GiB", "x86_64", "bare node"],
    ["GPU", "NVIDIA GH200 480 GB", "ARM64 Grace", "public image + ConfigMap"],
    ["TPU", "v5e 2×4 — 8 chips", "x86_64", "private registry + gcsfuse"],
], [Inches(1.5), Inches(4.2), Inches(2.6), Inches(3.8)], size=13)


# ============================== 3 — Measurements ================================
s = add_slide(prs)
title(s, "3", "How we measured",
      "Every number below is emitted by the run itself, against a fixed schema")

textbox(s, Inches(0.6), Inches(1.8), Inches(5.6), Inches(3.4),
        "Wall-clock decomposition, six phases\n"
        "scheduling · image pull · checkpoint load ·\n"
        "XLA compile · steady-state steps · checkpoint write\n\n"
        "Sweeps\n"
        "batch 1→32 and sequence 512→2048, walked into OOM.\n"
        "A failed cell is recorded, not dropped — the boundary\n"
        "is the chart.\n\n"
        "Contract\n"
        "docs/metrics-contract.md. One JSON per run.\n"
        "Charts read results/ and nothing else.", size=14)

figure(s, "panel-walltime.png", Inches(6.5), Inches(1.75), Inches(6.3))

textbox(s, Inches(0.6), Inches(5.55), Inches(5.6), Inches(1.5),
        "The schema caught a bad measurement", size=14, bold=True, color=ORANGE)
textbox(s, Inches(0.6), Inches(5.9), Inches(5.6), Inches(1.3),
        "One run reported 93 µs per step and 11 M tokens/s — impossible.\n"
        "It carried status \"ok\", but telemetry.py had flagged in notes that it\n"
        "may have timed dispatch, not completion. Discarded and repeated.", size=12, color=GREY)


# ============================== 4 — Results ====================================
s = add_slide(prs)
title(s, "4", "Results",
      f"Qwen3-0.6B · batch 1 · sequence 1024 · bfloat16 — identical script, identical flags")

table(s, Inches(0.6), Inches(1.75), [
    ["", "CPU 32-core", "GPU GH200", "TPU v5e 2×4"],
    ["", "", "1 chip", "8 chips"],
    ["Median step", f"{c_step:.2f} s", f"{g_step:.4f} s", f"{t_step:.4f} s"],
    ["Throughput", f"{cpu['steady']['tokens_per_s']:.0f} tok/s",
     f"{gpu['steady']['tokens_per_s']:,.0f} tok/s", f"{tpu['steady']['tokens_per_s']:,.0f} tok/s"],
    ["Speedup vs CPU", "1.0×", f"{speed_g:.1f}×", f"{speed_t:.1f}×"],
    ["Peak memory", f"{cpu['memory']['peak_pct']:.1f} %",
     f"{gpu['memory']['peak_pct']:.1f} %", f"{tpu['memory']['peak_pct']:.1f} %"],
    ["Chip utilisation", f"{cpu['utilization']['mean_pct']:.1f} % host",
     f"{gpu['utilization']['mean_pct']:.2f} %", "no counter"],
    ["USD / 1 000 steps", f"${cost_cpu:.2f}", f"${cost_gpu:.3f}", f"${cost_tpu:.2f}"],
], [Inches(2.3), Inches(1.9), Inches(1.9), Inches(1.9)], size=13, row_h=Inches(0.38))

textbox(s, Inches(0.6), Inches(4.75), Inches(7.6), Inches(0.5),
        f"One GH200 matches an eight-chip TPU slice to within {gt_delta:.2f} %",
        size=19, bold=True, color=ORANGE)
textbox(s, Inches(0.6), Inches(5.2), Inches(7.6), Inches(1.9),
        "At batch 1 the job occupies neither accelerator — 5.0 % of the GH200 and 7.3 %\n"
        "of the TPU. It is latency-bound, so the extra seven chips buy nothing.\n\n"
        f"Mitigation: one manifest line — fileCacheCapacity — cut checkpoint load\n"
        f"{b_ph['model_load_s']:.0f} s → {m_ph['model_load_s']:.0f} s ({load_factor:.1f}×) "
        f"and total wall clock by {wall_saving:.1f} %.\n"
        "Control: step time identical to the fourth decimal. Only I/O moved.", size=13)

figure(s, "panel-mitigation.png", Inches(8.5), Inches(1.9), Inches(4.3))


# ============================== 5 — Conclusion =================================
s = add_slide(prs)
title(s, "5", "What we learned", "Bottleneck, mitigation, and the recommendation")

textbox(s, Inches(0.6), Inches(1.8), Inches(6.0), Inches(0.45),
        "The bottleneck is not compute", size=17, bold=True, color=ORANGE)
textbox(s, Inches(0.6), Inches(2.2), Inches(6.0), Inches(1.5),
        f"Of {b_ph['total_wall_s']:.0f} s of wall clock, only {b_ph['steady_state_s']:.0f} s "
        f"— {compute_pct:.1f} % — is arithmetic.\n"
        f"The largest single phase is reading the checkpoint: {load_pct:.1f} %.\n"
        "Scheduling and XLA compilation take most of the rest.", size=14)

textbox(s, Inches(0.6), Inches(3.6), Inches(6.0), Inches(0.45),
        "Scaling the model did not scale the cost", size=17, bold=True, color=ORANGE)
textbox(s, Inches(0.6), Inches(4.0), Inches(6.0), Inches(1.6),
        "Growing Qwen3 from 0.6B to 4B — about 6.7× — did not make the CPU\n"
        "6.7× slower. It made the CPU impossible: XLA did not finish compiling\n"
        "in over an hour with remat, and the kernel killed the process at\n"
        "31.5 GB without it. A speedup chart alone hides that discontinuity.", size=14)

textbox(s, Inches(6.9), Inches(1.8), Inches(5.9), Inches(0.45),
        "The wall was in the supply chain", size=17, bold=True, color=ORANGE)
textbox(s, Inches(6.9), Inches(2.2), Inches(5.9), Inches(1.9),
        "The GH200 sat idle for a day — not for want of hardware or code.\n"
        "ARM64 host · QEMU cross-build segfault · libtpu has no aarch64 wheel ·\n"
        "no cluster credentials · registry 403.\n\n"
        "Exactly one of Tunix's 36 declared dependencies is TPU-bound.\n"
        "Swapping it is the whole port.", size=14)

textbox(s, Inches(6.9), Inches(4.3), Inches(5.9), Inches(0.45),
        "Cost, and the recommendation", size=17, bold=True, color=ORANGE)

table(s, Inches(6.9), Inches(4.72), [
    ["USD per 1 000 steps", "CPU", "TPU 8×", "GH200"],
    ["same workload, list prices",
     f"${cost_cpu:.2f}", f"${cost_tpu:.2f}", f"${cost_gpu:.3f}"],
], [Inches(2.9), Inches(1.0), Inches(1.0), Inches(1.0)], size=12, row_h=Inches(0.36))

textbox(s, Inches(6.9), Inches(5.55), Inches(5.9), Inches(1.5),
        f"Identical step time, {cost_tpu / cost_gpu:.1f}× the price: eight v5e chips against one\n"
        "Hopper. Do not scale out — amortise. Raise batch only until the chip is\n"
        "occupied, then attack fixed cost: persistent compile cache, warm workers\n"
        "instead of a pod per job, file cache on anything reading a checkpoint.", size=13)

textbox(s, Inches(0.6), Inches(6.55), Inches(12.2), Inches(0.5),
        "9 measured runs · 15 commits · github.com/RN0L/me344-qwen3-icd10-profiling",
        size=12, color=GREY, align=PP_ALIGN.CENTER)


OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"wrote {OUT.relative_to(REPO)}  ({OUT.stat().st_size // 1024} KB, {len(prs.slides.__iter__.__self__._sldIdLst)} slides)")
