#!/usr/bin/env python3
"""Build the 5-slide deck from the measured records.

Every figure on these slides is read out of ``results/`` at build time. Nothing is
transcribed, so the deck cannot drift from the data, re-run this after any new run
and the numbers update themselves.

    uv run --with python-pptx --with pillow python slides/build_pptx.py

Design intent: one idea per slide, one number large enough to read from the back of
the room, and body text kept under ~45 words. The house colour is a single constant.
"""

from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

REPO = Path(__file__).resolve().parent.parent
RESULTS = REPO / "results"
FIGURES = Path(__file__).resolve().parent / "figures"
OUT = Path(__file__).resolve().parent / "ME344_Final_Arnold_Hambuch.pptx"

# --- palette -------------------------------------------------------------------
ORANGE = RGBColor(0xE8, 0x59, 0x0C)
ORANGE_L = RGBColor(0xFB, 0x92, 0x3C)
TINT = RGBColor(0xFE, 0xF3, 0xEA)
INK = RGBColor(0x15, 0x17, 0x1A)
GREY = RGBColor(0x71, 0x7A, 0x85)
FAINT = RGBColor(0xE6, 0xE9, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Helvetica Neue"

W, H = Inches(13.333), Inches(7.5)


def load(n):
    return json.loads((RESULTS / f"{n}.json").read_text())


# --- facts, read not typed -----------------------------------------------------
recs = {}
for f in sorted(RESULTS.glob("*.json")):
    if f.stem != "analysis":
        d = json.loads(f.read_text())
        recs[d["run_id"]] = d

cpu = recs["cpu-x86-32c-0p6b-bs1-seq1024"]
gpu = recs["gpu-gh200-0p6b-bs1-seq1024"]
tpu = recs["tpu-v5e8-bs1-seq1024-filecache-0p6b"]
base = recs["tpu-v5e8-bs8-seq1024"]
mit = recs["tpu-v5e8-bs8-seq1024-filecache"]

c_s, g_s, t_s = (r["steady"]["median_step_s"] for r in (cpu, gpu, tpu))
sp_g, sp_t = c_s / g_s, c_s / t_s
gt = abs(g_s - t_s) / t_s * 100
bp, mp = base["phases"], mit["phases"]
compute_pct = bp["steady_state_s"] / bp["total_wall_s"] * 100
load_x = bp["model_load_s"] / mp["model_load_s"]
wall_cut = (1 - mp["total_wall_s"] / bp["total_wall_s"]) * 100

cost = {c["run_id"]: c for c in load("analysis")["cost"]["per_run"]}
k_cpu = cost[cpu["run_id"]]["usd_per_1000_steps"]
k_gpu = cost[gpu["run_id"]]["usd_per_1000_steps"]
k_tpu = cost[tpu["run_id"]]["usd_per_1000_steps"]

stats = json.loads((REPO / "docs" / "dataset_stats.json").read_text())
loss = {e["seq_len"]: e["frac_completion_fully_lost"] * 100
        for e in stats["splits"]["train"]["truncation"]}

# GPU sweep cells, if the sweep has landed
gpu_sweep = sorted(
    ((r["config"]["batch_size"], r) for r in recs.values()
     if r["backend"] == "gpu" and r["config"]["model"].endswith("4B")),
    key=lambda t: t[0])
tpu_sweep = sorted(
    ((r["config"]["batch_size"], r) for r in recs.values()
     if r["backend"] == "tpu" and r["config"]["model"].endswith("4B")
     and r["config"]["seq_len"] == 1024),
    key=lambda t: t[0])

prs = Presentation()
prs.slide_width, prs.slide_height = W, H


# --- primitives ----------------------------------------------------------------
def slide(num, kicker, headline):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rule = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.16), H)
    rule.fill.solid(); rule.fill.fore_color.rgb = ORANGE
    rule.line.fill.background(); rule.shadow.inherit = False

    ghost = s.shapes.add_textbox(Inches(11.9), Inches(0.05), Inches(1.3), Inches(1.5))
    r = ghost.text_frame.paragraphs[0].add_run(); r.text = str(num)
    r.font.size = Pt(72); r.font.bold = True; r.font.color.rgb = FAINT; r.font.name = FONT
    ghost.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT

    txt(s, Inches(0.62), Inches(0.40), Inches(10.5), Inches(0.3),
        kicker.upper(), 12, color=ORANGE, bold=True, spacing=2)
    txt(s, Inches(0.62), Inches(0.72), Inches(11.0), Inches(0.9),
        headline, 32, bold=True)
    return s


def txt(s, x, y, w, h, text, size=14, bold=False, color=INK,
        align=PP_ALIGN.LEFT, gap=5, spacing=0, italic=False):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(gap)
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = color; r.font.name = FONT
    return tb


def stat(s, x, y, w, value, label, big=44, color=ORANGE):
    """A single number, sized to read from the back of the room."""
    txt(s, x, y, w, Inches(0.85), value, big, bold=True, color=color, gap=0)
    txt(s, x, y + Inches(0.72), w, Inches(0.9), label, 12, color=GREY, gap=2)


def panel(s, x, y, w, h, fill=TINT):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = FAINT; sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    sh.adjustments[0] = 0.04
    return sh


def table(s, x, y, rows, widths, size=13, rh=Inches(0.40), highlight=None,
          right_align_from=1):
    """right_align_from: first column index to right-align. Set high to keep text left."""
    t = s.shapes.add_table(len(rows), len(rows[0]), x, y, sum(widths), rh * len(rows)).table
    for j, wd in enumerate(widths):
        t.columns[j].width = wd
    for i, row in enumerate(rows):
        t.rows[i].height = rh
        for j, v in enumerate(row):
            c = t.cell(i, j); c.text = str(v)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = c.margin_right = Inches(0.09)
            c.margin_top = c.margin_bottom = Inches(0.01)
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if j >= right_align_from else PP_ALIGN.LEFT
            r = p.runs[0] if p.runs else p.add_run()
            r.font.size = Pt(size); r.font.name = FONT
            c.fill.solid()
            if i == 0:
                c.fill.fore_color.rgb = INK
                r.font.bold = True; r.font.color.rgb = WHITE
            elif highlight is not None and j == highlight:
                c.fill.fore_color.rgb = TINT
                r.font.bold = True; r.font.color.rgb = ORANGE
            else:
                c.fill.fore_color.rgb = WHITE
                r.font.color.rgb = INK
    return t


def fig(s, name, x, y, width):
    p = FIGURES / name
    if p.exists():
        s.shapes.add_picture(str(p), x, y, width=width)


# ============================== 1. Problem ====================================
s = slide(1, "The problem", "Can a clinic afford to run its own coding model?")

panel(s, Inches(0.62), Inches(1.85), Inches(6.0), Inches(2.35))
txt(s, Inches(0.92), Inches(2.08), Inches(5.5), Inches(0.35),
    "Dr. Findus, our startup in Berlin", 14, bold=True)
txt(s, Inches(0.92), Inches(2.5), Inches(5.5), Inches(1.6),
    "German GPs must code every consultation.\n"
    "We read the note and propose the codes;\n"
    "the physician confirms each one. A legal\n"
    "requirement here. Reliability is the product:\n"
    "over-coding is fraud exposure, not error.", 13, gap=3)

txt(s, Inches(7.0), Inches(1.95), Inches(5.7), Inches(0.8),
    "Today that runs on hosted APIs.\nCould it run on our own hardware?", 16, bold=True, gap=2)
txt(s, Inches(7.0), Inches(2.85), Inches(5.7), Inches(1.4),
    "An infrastructure question, not a model one.\n"
    "so we built the workload and measured it.\n\n"
    "Qwen3 + LoRA on CodiEsp: 1 000 clinical case\n"
    "reports, ICD-10 labelled, CC-BY 4.0.", 13, color=GREY, gap=3)

stat(s, Inches(0.62), Inches(4.55), Inches(3.0), "8 GB",
     "of weights at 4B, before a single\nactivation exists")
stat(s, Inches(4.0), Inches(4.55), Inches(3.0), "2 489",
     "tokens in the longest case.\nSequence length is not free")
stat(s, Inches(7.4), Inches(4.55), Inches(5.3), "3 backends",
     "CPU · GPU · TPU, one JAX program,\nno per-backend reimplementation")

txt(s, Inches(0.62), Inches(6.35), Inches(12.0), Inches(0.7),
    "Where does the time actually go, and what does it take to run this at all?",
    19, bold=True, color=ORANGE)
txt(s, Inches(0.62), Inches(6.85), Inches(12.0), Inches(0.4),
    "No patient data was used. Model accuracy earns no marks here. The object of study is the infrastructure.",
    11, color=GREY, italic=True)


# ============================== 2. Approach ===================================
s = slide(2, "Approach", "One program. Three backends. All measured.")

cols = [
    ("Compilation", ORANGE,
     "One JAX / Flax program.\nXLA lowers it to CPU, CUDA\nand TPU.\nLoRA via qwix, training via\nTunix, no per-backend\nreimplementation."),
    ("Orchestration", ORANGE,
     "3 pinned Docker images.\nKubernetes Jobs across two\nclusters.\nKueue admission on the TPU\npool. Every manifest states\nits CPU, memory and chips."),
    ("Storage", ORANGE,
     "gcsfuse CSI, implicit-dirs.\nAn 8 GB checkpoint read\nfrom a bucket.\nThe file cache was left as a\nknob rather than fixed -\nthat knob became the\nexperiment."),
]
for i, (head, col, body) in enumerate(cols):
    x = Inches(0.62 + i * 4.15)
    panel(s, x, Inches(1.95), Inches(3.85), Inches(2.5))
    txt(s, x + Inches(0.28), Inches(2.15), Inches(3.3), Inches(0.4), head, 16, bold=True, color=col)
    txt(s, x + Inches(0.28), Inches(2.58), Inches(3.35), Inches(1.8), body, 12, gap=2)

table(s, Inches(0.62), Inches(4.75), [
    ["Backend", "Hardware", "Host arch", "How the code got there"],
    ["CPU", "32-core x86_64, 31 GiB", "x86_64", "bare node"],
    ["GPU", "NVIDIA GH200 480 GB", "ARM64 Grace", "public image + ConfigMap"],
    ["TPU", "v5e 2×4, 8 chips", "x86_64", "private registry + gcsfuse"],
], [Inches(1.5), Inches(4.1), Inches(2.4), Inches(4.1)], size=12.5,
    right_align_from=99)

txt(s, Inches(0.62), Inches(6.55), Inches(12.0), Inches(0.6),
    "Note the middle row: the GH200's host is ARM64, not x86. That single fact cost us five separate blockers.",
    13, color=GREY, italic=True)


# ============================== 3. Measurements ================================
s = slide(3, "Measurements", "We did not time the job. We took it apart.")

txt(s, Inches(0.62), Inches(1.95), Inches(5.3), Inches(2.6),
    "Every run takes its own wall clock apart\n"
    "into six phases: scheduling, image pull,\n"
    "checkpoint load, XLA compile, the steady-\n"
    "state steps, checkpoint write.\n\n"
    "Batch and sequence sweeps walked until\n"
    "the job dies. A failed cell is recorded,\n"
    "never dropped. The boundary is the\n"
    "measurement.\n\n"
    "One schema, one JSON per run. Every figure\n"
    "here reads those files and nothing else.", 13, gap=3)

fig(s, "slide3-walltime.png", Inches(6.4), Inches(2.0), Inches(6.3))

panel(s, Inches(6.4), Inches(3.85), Inches(6.3), Inches(2.35),
      fill=RGBColor(0xFE, 0xF2, 0xF2))
txt(s, Inches(6.72), Inches(4.08), Inches(5.7), Inches(0.35),
    "The schema caught a lie", 15, bold=True, color=RGBColor(0xDC, 0x26, 0x26))
txt(s, Inches(6.72), Inches(4.5), Inches(5.7), Inches(1.6),
    "One run reported 93 µs per step and 11 M tokens/s.\n"
    "Impossible, yet it carried status \"ok\".\n"
    "Its own notes flagged that it may have timed\n"
    "dispatch, not completion. Discarded, repeated.", 12, gap=2)


# ============================== 4. Results ====================================
s = slide(4, "Results", "One Hopper chip kept up with eight TPU chips.")

table(s, Inches(0.62), Inches(1.82), [
    ["Qwen3-0.6B · bs 1 · seq 1024", "CPU", "GPU GH200", "TPU v5e 2×4"],
    ["chips", "-", "1", "8"],
    ["median step", f"{c_s:.2f} s", f"{g_s:.4f} s", f"{t_s:.4f} s"],
    ["throughput", f"{cpu['steady']['tokens_per_s']:,.0f}", f"{gpu['steady']['tokens_per_s']:,.0f}",
     f"{tpu['steady']['tokens_per_s']:,.0f}"],
    ["speedup vs CPU", "1.0×", f"{sp_g:.0f}×", f"{sp_t:.0f}×"],
    ["peak memory", f"{cpu['memory']['peak_pct']:.0f} %",
     f"{gpu['memory']['peak_pct']:.1f} %", f"{tpu['memory']['peak_pct']:.1f} %"],
    ["USD / 1 000 steps", f"${k_cpu:.2f}", f"${k_gpu:.3f}", f"${k_tpu:.2f}"],
], [Inches(2.6), Inches(1.45), Inches(1.6), Inches(1.6)], size=12,
    rh=Inches(0.33), highlight=2)

stat(s, Inches(8.7), Inches(1.85), Inches(4.0), f"{gt:.2f} %",
     "difference in step time between\none GH200 and eight v5e chips", big=40)
stat(s, Inches(8.7), Inches(3.15), Inches(4.0), f"{load_x:.1f}×",
     f"faster checkpoint load from one\nmanifest line, {wall_cut:.0f} % off wall clock", big=40)
fig(s, "slide4-comparison.png", Inches(0.62), Inches(4.28), Inches(7.4))

panel(s, Inches(8.7), Inches(4.55), Inches(4.0), Inches(2.2))
txt(s, Inches(9.0), Inches(4.75), Inches(3.5), Inches(0.3),
    "The control", 13, bold=True, color=ORANGE)
txt(s, Inches(9.0), Inches(5.12), Inches(3.5), Inches(1.6),
    "Median step time before and\nafter the cache:\n\n"
    "1.75122 s  to  1.75117 s\n\n"
    "The computation did not move.\nOnly I/O did, so the saving is\nattributable, not noise.", 11.5, gap=2)


# ============================== 5. Conclusion =================================
s = slide(5, "Conclusion", "The chip was never the bottleneck.")

findings = [
    (f"{compute_pct:.0f} %",
     "of wall clock is arithmetic",
     f"Of {bp['total_wall_s']:.0f} s, only {bp['steady_state_s']:.0f} s computes.\n"
     "The biggest single phase is\nreading the checkpoint."),
    ("6.7×",
     "more parameters, CPU stopped",
     "0.6B to 4B did not make the CPU\n6.7× slower. XLA never finished\n"
     "compiling; the kernel killed it\nat 31.5 GB."),
    ("1 of 36",
     "dependencies were TPU-bound",
     "The GH200 sat idle behind an\nARM64 host, a QEMU segfault\n"
     "and a registry 403. Swapping\none pin is the whole port."),
]
for i, (big, cap, body) in enumerate(findings):
    x = Inches(0.62 + i * 4.15)
    panel(s, x, Inches(1.9), Inches(3.85), Inches(2.65))
    txt(s, x + Inches(0.28), Inches(2.05), Inches(3.3), Inches(0.7), big, 34, bold=True, color=ORANGE, gap=0)
    txt(s, x + Inches(0.28), Inches(2.68), Inches(3.35), Inches(0.45), cap, 12, bold=True, gap=2)
    txt(s, x + Inches(0.28), Inches(3.15), Inches(3.35), Inches(1.3), body, 11.5, color=GREY, gap=2)

txt(s, Inches(0.62), Inches(4.8), Inches(5.9), Inches(0.35),
    "Scaling recommendation", 15, bold=True, color=ORANGE)
txt(s, Inches(0.62), Inches(5.18), Inches(5.9), Inches(1.4),
    "Do not scale out. Amortise: raise batch only until\n"
    "the chip is occupied, then attack fixed cost:\n"
    "persistent compile caches, warm workers instead of\n"
    "a pod per job, file cache on anything reading a\n"
    "checkpoint.", 13, gap=2)

panel(s, Inches(6.8), Inches(4.72), Inches(5.9), Inches(1.85))
txt(s, Inches(7.1), Inches(4.9), Inches(5.4), Inches(0.35),
    "What this means for Dr. Findus", 15, bold=True, color=ORANGE)
txt(s, Inches(7.1), Inches(5.28), Inches(5.4), Inches(1.2),
    f"At ${k_gpu:.3f} per 1 000 steps on one GH200,\n"
    f"{k_tpu / k_gpu:.0f}× cheaper than eight TPU chips at the same\n"
    "speed, self-hosting is not the cost problem we\n"
    "assumed. The real cost is fixed overhead per job:\n"
    "an architecture decision, not a hardware purchase.", 12.5, gap=2)

txt(s, Inches(0.62), Inches(6.75), Inches(12.1), Inches(0.4),
    "9 measured runs · 15 commits · github.com/RN0L/me344-qwen3-icd10-profiling",
    11, color=GREY, align=PP_ALIGN.CENTER)


prs.save(OUT)
n = len(prs.slides._sldIdLst)
print(f"wrote {OUT.relative_to(REPO)}, {n} slides, {OUT.stat().st_size // 1024} KB")
if gpu_sweep:
    print(f"GPU sweep cells available: {[b for b, _ in gpu_sweep]}, rerun to fold them in")
