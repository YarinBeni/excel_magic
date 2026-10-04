"""Round-2 result figures (PNG) from the run directories.

  python docs/research_3models/make_figures.py      (from the repo root; writes docs/research_3models/figures/*.png)
"""
from __future__ import annotations

import json
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402

RUNS = Path("reports/runs")
OUT = Path("docs/research_3models/figures")
INK, MUTED, GRID = "#1f2430", "#6b7280", "#e5e7eb"
BASE, GOOD, BAD, NEU = "#9aa3b2", "#2f6db5", "#c2553a", "#8a6fb0"
plt.rcParams.update({"font.size": 10, "axes.edgecolor": GRID, "axes.labelcolor": INK, "xtick.color": MUTED,
                     "ytick.color": MUTED, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.6, "axes.axisbelow": True,
                     "figure.dpi": 150, "savefig.bbox": "tight"})


def rows(pattern: str, name: str) -> pd.DataFrame:
    out = []
    for d in sorted(RUNS.glob(pattern)):
        f = d / name
        if f.exists():
            out += [dict(json.loads(x), _run=d.name) for x in f.open()]
    return pd.DataFrame(out)


def paired(df, a, b):
    p = df.pivot_table(index="qid", columns="config", values="correct", aggfunc="first").astype(float)
    d = (p[b] - p[a]).dropna()
    return d.mean(), d.std(ddof=1) / np.sqrt(len(d))


def hbar(ax, labels, vals, errs, colors, ref=None, fmt="{:.3f}", xlim=None):
    y = np.arange(len(labels))[::-1]
    ax.barh(y, vals, xerr=errs, color=colors, height=0.6, error_kw={"ecolor": MUTED, "lw": 1, "capsize": 2})
    ax.set_yticks(y, labels)
    ax.grid(axis="y", visible=False)
    for yi, v, e in zip(y, vals, errs):
        ax.text(v + e + ((xlim[1] - xlim[0]) * 0.01 if xlim else 0), yi, fmt.format(v), va="center", fontsize=9, color=INK)
    if ref is not None:
        ax.axvline(ref, color=MUTED, lw=1, ls="--")
    if xlim:
        ax.set_xlim(*xlim)


def fig_bird_ladder(h2):
    d = h2[(h2.model == "qwen3coder30b") & (h2.split == "dev")].drop_duplicates(["qid", "config"], keep="last")
    order = [("A", "A  plain LLM"), ("K1", "+ column descriptions"), ("K2", "+ data profile"),
             ("K3", "+ rule card"), ("G", "+ deterministic checks (with rule card)"),
             ("GD", "+ GLiClass doubt signal"), ("GN", "GN: descriptions + profile + checks"),
             ("SC8N", "8 answers, majority vote"), ("GN8", "GN + 8 answers, vote among passing")]
    labs, vals, errs, cols = [], [], [], []
    for c, lab in order:
        g = d[d.config == c]
        labs.append(lab)
        vals.append(g.correct.mean())
        errs.append(0 if c == "A" else paired(d, "A", c)[1])
        cols.append(BASE if c == "A" else BAD if c == "K3" else GOOD if c in ("G", "GN", "GN8", "GD") else NEU)
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    hbar(ax, labs, vals, errs, cols, ref=vals[0], xlim=(0.6, 0.78))
    ax.set_xlabel("execution accuracy (373 design questions); bars = 1 s.e. of the paired difference to A")
    ax.set_title("BIRD: what each harness step adds (Qwen3-Coder-30B, zero-shot)", loc="left", color=INK, fontsize=11)
    fig.savefig(OUT / "1_bird_ladder.png")
    plt.close(fig)


def fig_bird_models(h2):
    d = h2[h2.split == "dev"].drop_duplicates(["qid", "config", "model"], keep="last")
    names = {"qwen3coder30b": "Qwen3-Coder-30B", "gptoss120b": "gpt-oss-120b", "glm45air": "GLM-4.5-Air"}
    fig, ax = plt.subplots(figsize=(7, 2.8))
    for i, (m, lab) in enumerate(names.items()):
        g = d[d.model == m]
        a, gn = g[g.config == "A"].correct.mean(), g[g.config == "GN"].correct.mean()
        diff, se = paired(g, "A", "GN")
        y = len(names) - 1 - i
        ax.plot([a, gn], [y, y], color=GRID, lw=3, zorder=1)
        ax.scatter([a], [y], color=BASE, s=60, zorder=2, label="plain LLM" if i == 0 else None)
        ax.scatter([gn], [y], color=GOOD, s=60, zorder=2, label="same LLM + harness (GN)" if i == 0 else None)
        ax.text(gn + 0.004, y, f"+{diff:.3f} (se {se:.3f})", va="center", fontsize=9, color=INK)
    ax.set_yticks(range(len(names))[::-1], list(names.values()))
    ax.grid(axis="y", visible=False)
    ax.set_xlim(0.6, 0.78)
    ax.set_xlabel("execution accuracy, 373 design questions")
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.25), ncol=2, fontsize=9)
    ax.set_title("BIRD: the harness gain holds across three open LLMs", loc="left", color=INK, fontsize=11)
    fig.savefig(OUT / "2_bird_models.png")
    plt.close(fig)


def fig_bird_heldout(h2):
    d = h2[(h2.split == "heldout") & (h2.model == "qwen3coder30b")]
    order = [("A", "A  plain LLM"), ("SC8N", "8 answers, majority vote"), ("GN", "GN: harness"),
             ("GN8", "GN + 8 answers"), ("GNL", "GN + fine-tuned GLiClass column finder"),
             ("GNLz", "GN + zero-shot GLiClass column finder")]
    first = d.drop_duplicates(["qid", "config"], keep="first")
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 3.4), gridspec_kw={"width_ratios": [1.6, 1]}, sharey=True)
    labs, vals, errs, cols, chars = [], [], [], [], []
    for c, lab in order:
        g = first[first.config == c]
        labs.append(lab)
        vals.append(g.correct.mean())
        errs.append(0 if c == "A" else paired(first, "A", c)[1])
        chars.append(g.schema_chars.mean())
        cols.append(BASE if c == "A" else BAD if c == "GNLz" else GOOD if c.startswith("GN") else NEU)
    hbar(ax1, labs, vals, errs, cols, ref=vals[0], xlim=(0.5, 0.8))
    ax1.set_xlabel("execution accuracy, 125 held-out questions (scored once)")
    hbar(ax2, labs, chars, [0] * len(chars), cols, fmt="{:,.0f}", xlim=(0, 14000))
    ax2.set_xlabel("prompt size, characters per question")
    fig.suptitle("BIRD held-out: a fine-tuned GLiClass keeps accuracy with a 2.6x smaller prompt",
                 x=0.01, ha="left", color=INK, fontsize=11)
    fig.savefig(OUT / "3_bird_heldout.png")
    plt.close(fig)


def fig_gli():
    m = json.load(open(sorted(RUNS.glob("*V9_gliclass_ft"))[-1] / "metrics.json"))
    c = json.load(open(sorted(RUNS.glob("*V10_gli_compare"))[-1] / "metrics.json"))
    fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(11, 3.3))
    z, t = m["linker"]["zero_shot"], m["linker"]["tuned"]
    x = np.arange(3)
    ax1.bar(x - 0.18, [z["recall@10"], z["all@20"], m["checker"]["auroc"]["sig_gliclass"]], 0.34, color=BASE, label="zero-shot")
    ax1.bar(x + 0.18, [t["recall@10"], t["all@20"], m["checker"]["auroc"]["tuned"]], 0.34, color=GOOD, label="fine-tuned on design questions")
    ax1.axhline(m["checker"]["auroc"]["sig_judge"], xmin=0.7, xmax=0.98, color=NEU, ls="--", lw=1)
    ax1.text(2.0, m["checker"]["auroc"]["sig_judge"] + 0.015, "LLM judge", ha="center", fontsize=8, color=NEU)
    ax1.set_xticks(x, ["found\n(top 10)", "all found\n(top 20)", "checker\nAUROC"])
    ax1.set_ylim(0, 1.0)
    ax1.legend(frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=1)
    ax1.grid(axis="x", visible=False)
    ax1.set_title("GLiClass, held-out questions", loc="left", fontsize=10, color=INK)
    names = {"knowledgator/gliclass-large-v3.0": "GLiClass large", "fastino/gliner2-large-v1": "GLiNER2 large",
             "fastino/GLiNER2.5-Decide": "GLiNER2.5-Decide"}
    ok = [k for k in names if "error" not in c.get(k, {"error": 1})]
    ax2.barh([names[k] for k in ok][::-1], [c[k]["linker_recall@10"] for k in ok][::-1], color=[BASE, BASE, GOOD][-len(ok):][::-1], height=0.55)
    for i, k in enumerate(ok[::-1]):
        ax2.text(c[k]["linker_recall@10"] + 0.01, i, f"{c[k]['linker_recall@10']:.3f}", va="center", fontsize=9)
    ax2.set_xlim(0, 1)
    ax2.grid(axis="y", visible=False)
    ax2.set_title("zero-shot column finder (top 10)", loc="left", fontsize=10, color=INK)
    ax3.barh([names[k] for k in ok][::-1], [max(c[k]["checker_auroc_pos"], c[k]["checker_auroc_diff"]) for k in ok][::-1],
             color=BASE, height=0.55)
    for i, k in enumerate(ok[::-1]):
        v = max(c[k]["checker_auroc_pos"], c[k]["checker_auroc_diff"])
        ax3.text(v + 0.01, i, f"{v:.3f}", va="center", fontsize=9)
    ax3.axvline(0.5, color=MUTED, ls="--", lw=1)
    ax3.set_xlim(0.4, 0.8)
    ax3.grid(axis="y", visible=False)
    ax3.set_title("zero-shot answer checker (AUROC)", loc="left", fontsize=10, color=INK)
    fig.suptitle("Small classifiers: zero-shot is weak; fine-tuning makes the column finder near-perfect",
                 x=0.01, ha="left", color=INK, fontsize=11)
    fig.tight_layout()
    fig.savefig(OUT / "4_gliclass.png")
    plt.close(fig)


def fig_relbench():
    cap = rows("*V4_relbench_capability", "capability_rows.jsonl")
    cap = cap[cap._run == cap.groupby("_run").size().idxmax()]
    hyp = rows("*V8_relbench_hypo", "hypo_rows.jsonl")
    hyp = hyp[hyp._run == hyp.groupby("_run").size().idxmax()]
    short = {"rel-avito/user-clicks": "avito clicks", "rel-avito/user-visits": "avito visits",
             "rel-event/user-ignore": "event ignore", "rel-event/user-repeat": "event repeat",
             "rel-f1/driver-dnf": "f1 DNF", "rel-f1/driver-top3": "f1 top-3", "rel-hm/user-churn": "hm churn",
             "rel-trial/study-outcome": "trial outcome"}
    tasks = list(short)
    D = cap[cap.config == "D"].groupby("task").auroc.mean()
    H = hyp.groupby("task").auroc.mean()
    F = hyp.groupby("task").auroc_fm.mean()
    fig, ax = plt.subplots(figsize=(8, 4))
    y = np.arange(len(tasks))[::-1]
    for yi, t in zip(y, tasks):
        ax.plot([D[t], max(F[t], H[t])], [yi, yi], color=GRID, lw=3, zorder=1)
    ax.scatter([D[t] for t in tasks], y, color=BASE, s=45, zorder=2, label="LLM alone, SQL only")
    ax.scatter([F[t] for t in tasks], y, color=NEU, s=45, zorder=2, marker="s", label="frozen relational model")
    ax.scatter([H[t] for t in tasks], y, color=GOOD, s=45, zorder=3, marker="D", label="frozen model + LLM hypotheses")
    ax.set_yticks(y, [short[t] for t in tasks])
    ax.grid(axis="y", visible=False)
    ax.axvline(0.5, color=MUTED, ls="--", lw=1)
    ax.set_xlim(0.45, 0.97)
    ax.set_xlabel("official test AUROC (0.5 = chance)")
    ax.legend(frameon=False, fontsize=8, loc="lower right")
    ax.set_title(f"RelBench: the models add a skill the LLM lacks; hypotheses add a little "
                 f"(mean {D[tasks].mean():.3f} / {F[tasks].mean():.3f} / {H[tasks].mean():.3f})",
                 loc="left", color=INK, fontsize=10.5)
    fig.savefig(OUT / "5_relbench.png")
    plt.close(fig)


def fig_insight():
    s = rows("*V3_insightbench_all", "scores.jsonl")
    s = s[s._run >= "20261004"]
    labs = {"P": "data profile + checklist", "PC": "+ GLiClass coverage signal", "Q0": "question agenda",
            "Q": "agenda ranked by GLiClass"}
    vals, errs = [], []
    for c in labs:
        g = s[s.config == c]
        run = g._run.iloc[0]
        base = s[(s._run == run) & (s.config == "D")].set_index("flag").g_eval
        d = (g.set_index("flag").g_eval - base).dropna()
        vals.append(d.mean())
        errs.append(d.std(ddof=1) / np.sqrt(len(d)))
    fig, ax = plt.subplots(figsize=(7, 2.6))
    y = np.arange(len(labs))[::-1]
    ax.barh(y, vals, xerr=errs, color=BAD, height=0.55, error_kw={"ecolor": MUTED, "lw": 1, "capsize": 2})
    for yi, v in zip(y, vals):
        ax.text(0.003, yi, f"{v:+.3f}", va="center", ha="left", fontsize=9, color=INK)
    ax.set_yticks(y, list(labs.values()))
    ax.axvline(0, color=INK, lw=1)
    ax.set_xlim(-0.09, 0.02)
    ax.grid(axis="y", visible=False)
    ax.set_xlabel("change in insight score vs the plain agent (same run, 100 tables)")
    ax.set_title("InsightBench: nothing beat the plain agent (dropped)", loc="left", color=INK, fontsize=11)
    fig.savefig(OUT / "6_insightbench.png")
    plt.close(fig)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    h2 = rows("*V7_sql_harness2", "harness2_rows.jsonl")
    fig_bird_ladder(h2)
    fig_bird_models(h2)
    fig_bird_heldout(h2)
    fig_gli()
    fig_relbench()
    fig_insight()
    print(sorted(p.name for p in OUT.glob("*.png")))


if __name__ == "__main__":
    main()
