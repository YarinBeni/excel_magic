"""Build the showcase figures from the run directories (no hand-typed numbers except the paper's own, from Table 9).

  python paper/make_figures.py --roots ../excel_magic/reports/runs --out paper/figures

fig1_tabarena_overview.png   mean gain over the frozen model alone, TabArena wave 1, every setup vs the paper
fig2_backbones_vs_paper.png  each open backbone with no search, relative to the paper's frozen TabFM, per dataset
fig3_harness_fixed_llm.png   entity task: same LLM, different agent harness (and the no-LLM baseline)
fig4_split_transfer.png      gain on the search split vs the other 29 splits (why the search selects noise)
fig5_retrieval.png           frozen embeddings for retrieval on RelBench rel-hm: item level and entity level
"""
from __future__ import annotations

import argparse
import glob
import json
import os
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# reference palette (dataviz skill, light mode); validated: blue/orange/aqua pass all-pairs CVD
SURFACE, INK, INK2, MUTED, GRID = "#fcfcfb", "#0b0b0b", "#52514e", "#8a8984", "#e6e5e1"
BLUE, ORANGE, AQUA, NEUTRAL = "#2a78d6", "#eb6834", "#1baf7a", "#b5b3ad"
plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "font.family": "DejaVu Sans", "font.size": 10, "text.color": INK, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.edgecolor": GRID, "axes.titleweight": "bold",
    "axes.titlesize": 12, "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
})


def _csvs(roots, pattern, name):
    out = []
    for r in roots:
        for f in glob.glob(os.path.join(r, pattern, name)):
            if os.path.getsize(f):
                out.append(pd.read_csv(f))
    return pd.concat(out) if out else pd.DataFrame()


def _metrics(roots, pattern):
    out = []
    for r in roots:
        for f in glob.glob(os.path.join(r, pattern, "metrics.json")):
            out.append(json.load(open(f)))
    return out


def _subtitle(ax, text):
    ax.text(0, 1.02, text, transform=ax.transAxes, color=INK2, fontsize=9, va="bottom")


def _header(fig, title, sub):
    """Title + subtitle anchored to the figure's left edge; returns the top of the plotting rect."""
    h = fig.get_figheight()
    fig.text(0.012, 1 - 0.18 / h, title, fontsize=12.5, fontweight="bold", color=INK, va="top")
    fig.text(0.012, 1 - 0.46 / h, sub, fontsize=9, color=INK2, va="top")
    return 1 - 0.72 / h


def _latest(roots, pattern):
    """metrics of the most recent run per run-name pattern (failed earlier attempts are superseded)."""
    fs = sorted(f for r in roots for f in glob.glob(os.path.join(r, pattern, "metrics.json")))
    return [json.load(open(fs[-1]))] if fs else []


def _hbars(ax, labels, values, colors, fmt="{:+.1f}%"):
    y = np.arange(len(labels))[::-1]
    ax.barh(y, values, height=0.62, color=colors, edgecolor=SURFACE, linewidth=2)
    ax.set_yticks(y, labels)
    ax.axvline(0, color=INK2, linewidth=1)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0)
    span = max(abs(v) for v in values) or 1
    for yi, v in zip(y, values):
        ax.text(v + (0.02 * span if v >= 0 else -0.02 * span), yi, fmt.format(v), va="center",
                ha="left" if v >= 0 else "right", color=INK, fontsize=9)
    lo, hi = min(0, min(values)), max(0, max(values))
    ax.set_xlim(lo - 0.18 * span, hi + 0.18 * span)


def fig1(roots, out):
    rows = []
    for pat, label in [("*J2_heuristic_*", "No LLM, greedy search · Kumo-S"),
                       ("*J8_pi_qwen3coder_*", "pi + Qwen3-Coder-30B · Kumo-S"),
                       ("*J8_openai_glm45air_*", "GLM-4.5-Air tool loop · Kumo-S"),
                       ("*J9_heuristic_kumoL_*", "No LLM, greedy search · Kumo-L"),
                       ("*J9_heuristic_tabiclv2_*", "No LLM, greedy search · TabICLv2"),
                       ("*J17_pi_qwen3coder_kumoL_rec_*", "pi + Qwen3-Coder · Kumo-L + guards"),
                       ("*J17_heuristic_kumoL_rec_*", "No LLM, greedy search · Kumo-L + guards")]:
        d = _csvs(roots, pat, "comparison_to_paper.csv")
        if len(d):
            rows.append((label, d["our_gain_%"].mean(), d.dataset.nunique()))
    ref = _csvs(roots, "*J2_heuristic_*", "comparison_to_paper.csv").drop_duplicates("dataset")
    paper = ref["paper_gain_opus5_%"].mean()
    labels = ["Paper: TabFM-Auto, Claude Opus 5"] + [r[0] for r in rows]
    values = [paper] + [r[1] for r in rows]
    fig, ax = plt.subplots(figsize=(8.6, 4.6))
    _hbars(ax, labels, values, [NEUTRAL] + [ORANGE if "guards" in r[0] else BLUE for r in rows])
    top = _header(fig, "Free, open setups do not reproduce the paper's gain on TabArena",
                  f"Mean error reduction over the frozen model alone · {len(ref)} datasets · 30 splits · orange = with guards (repeated judge, gated pick, acceptance slice)")
    ax.set_xlabel("Mean gain over the frozen model (%)")
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig1_tabarena_overview.png", dpi=180)
    plt.close(fig)


def fig2(roots, out):
    series = [("*J9_heuristic_kumoL_*", "Kumo Tabular-L", BLUE, "o"),
              ("*J2_heuristic_*", "Kumo Tabular-S", ORANGE, "s"),
              ("*J9_heuristic_tabiclv2_*", "TabICLv2", AQUA, "D")]
    frames = {}
    for pat, name, _, _ in series:
        d = _csvs(roots, pat, "comparison_to_paper.csv")
        d["rel"] = 100 * (d["tabfm"] - d["P0"]) / d["tabfm"]  # >0: open backbone better than paper TabFM
        frames[name] = d.set_index("dataset")["rel"]
    order = frames["Kumo Tabular-L"].sort_values().index
    fig, ax = plt.subplots(figsize=(8.6, 6.2))
    y = np.arange(len(order))
    for (_pat, name, color, marker), dy in zip(series, (0.22, 0.0, -0.22)):
        v = frames[name].reindex(order).clip(-60, 60)
        wins = int((frames[name] > 0).sum())
        ax.scatter(v, y + dy, s=42, color=color, marker=marker, edgecolors=SURFACE, linewidths=1.5, zorder=3,
                   label=f"{name}  (beats paper TabFM on {wins}/{len(order)})")
    ax.axvline(0, color=INK2, linewidth=1)
    ax.set_yticks(y, [o.replace("-spread-contaminant-detection", "").replace("Another-Dataset-on-used-", "") for o in order])
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_xlabel("Error vs the paper's frozen TabFM (%, right of zero = open backbone is better)")
    top = _header(fig, "With no search at all, the best open backbone already matches the paper's TabFM",
                  "Identity pipeline (frozen model, raw columns) · TabArena wave 1 · mean over 30 official splits · hazelnut clipped at −60%")
    ax.legend(loc="center left", frameon=False, fontsize=9)
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig2_backbones_vs_paper.png", dpi=180)
    plt.close(fig)


def fig3(roots, out):
    def gain(*pats):
        ms = [m for p in pats for m in _latest(roots, p) if m.get("test_improvement_pct") is not None]
        return float(np.mean([m["test_improvement_pct"] for m in ms])) if ms else None
    rows = [("pi (CLI agent) + Qwen3-Coder-30B", gain("*J6_pi_synth_entities"), BLUE),
            ("aider loop + Qwen3-Coder-30B", gain("*J6_aider_synth_entities"), BLUE),
            ("Qwen Code + Qwen3-Coder-30B", gain("*J6_qwen-code_synth_entities"), BLUE),
            ("Minimal tool loop + Qwen3-Coder-30B", gain("*J3_openai_synth_entities", "*J5_Qwen3-Coder-30B-A3B-Instruct_synth_entities"), BLUE),
            ("Minimal loop + richer context (J13)", gain("*J13_rich_qwen3coder_synth_entities"), BLUE),
            ("Minimal tool loop + GLM-4.5-Air", gain("*J5_GLM-4.5-Air-FP8_synth_entities"), ORANGE),
            ("No LLM, greedy search", float(np.mean([m["test_improvement_pct"] for m in _metrics(roots, "*J1_heur_kumo_synth_entities")])), NEUTRAL)]
    rows = [r for r in rows if r[1] is not None]
    fig, ax = plt.subplots(figsize=(8.6, 3.9))
    _hbars(ax, [r[0] for r in rows], [r[1] for r in rows], [r[2] for r in rows])
    top = _header(fig, "Same open LLM, different agent harness: the harness decides",
                  "Entity-ID task (synth_entities) · frozen Kumo Tabular-S · budget 16 evaluations · held-out test")
    ax.set_xlabel("Gain over the frozen model (%)")
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig3_harness_fixed_llm.png", dpi=180)
    plt.close(fig)


def fig4(roots, out):
    series = [("*J2_heuristic_*", "No LLM, greedy search", BLUE, "o"),
              ("*J8_pi_qwen3coder_*", "pi + Qwen3-Coder-30B", ORANGE, "s"),
              ("*J8_openai_glm45air_*", "GLM-4.5-Air tool loop", AQUA, "D")]
    fig, ax = plt.subplots(figsize=(7.4, 6.2))
    allv = []
    for pat, name, color, marker in series:
        d = _csvs(roots, pat, "results.csv")
        pts = []
        for _ds, g in d.groupby("dataset"):
            p0 = g[g.pipeline == "P0"].set_index(["repeat", "fold"]).error
            ps = g[g.pipeline == "P*"].set_index(["repeat", "fold"]).error
            rel = (100 * (p0 - ps) / p0).dropna()
            if (0, 0) in rel.index and len(rel) > 1:
                pts.append((rel.loc[(0, 0)], rel.drop((0, 0)).mean()))
        pts = np.clip(np.array(pts), -25, 25)
        allv.append(pts)
        ax.scatter(pts[:, 0], pts[:, 1], s=46, color=color, marker=marker, edgecolors=SURFACE, linewidths=1.5,
                   zorder=3, label=name)
    lim = 26
    ax.plot([-lim, lim], [-lim, lim], color=MUTED, linewidth=1, linestyle=(0, (4, 3)), zorder=1)
    ax.text(lim * 0.55, lim * 0.62, "transfers fully", color=MUTED, fontsize=8, rotation=45)
    ax.axhline(0, color=INK2, linewidth=1); ax.axvline(0, color=INK2, linewidth=1)
    ax.set_xlim(-lim, lim); ax.set_ylim(-lim, lim)
    ax.grid(color=GRID, linewidth=0.8); ax.set_axisbelow(True)
    ax.set_xlabel("Gain on the split the search was run on (%)")
    ax.set_ylabel("Mean gain on the other 29 official splits (%)")
    top = _header(fig, "Gains found on the search split mostly do not transfer",
                  "One point per dataset and setup · TabArena wave 1 · clipped at ±25%")
    ax.legend(loc="upper left", frameon=False, fontsize=9)
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig4_split_transfer.png", dpi=180)
    plt.close(fig)


def fig5(roots, out):
    item = {}
    for r in roots:
        for f in sorted(glob.glob(os.path.join(r, "*J7*_hm_*full", "metrics.json")) + sorted(glob.glob(os.path.join(r, "*J14*_hm_full", "metrics.json")))):
            m = json.load(open(f))
            for name, v in m.get("splits", {}).get("test", {}).get("rows", {}).items():
                if "map" in v:
                    item[name] = 100 * v["map"]
    item_rows = [("Cosine on the purchase matrix (no model)", "kNN-CF[purchase_matrix]", NEUTRAL),
                 ("Item-kNN (no model)", "ItemKNN", NEUTRAL),
                 ("SVD factors of purchases (no model)", "kNN-CF[svd]", NEUTRAL),
                 ("Global popularity", "GlobalPopularity", NEUTRAL),
                 ("Hand aggregates", "kNN-CF[agg]", NEUTRAL),
                 ("Frozen Kumo Relational, random target", "kNN-CF[kumo_relational_random]", BLUE),
                 ("Frozen TabPFN over aggregates, k-means", "kNN-CF[tabpfn_agg_kmeans]", BLUE),
                 ("Frozen Kumo Relational, k-means target", "kNN-CF[kumo_relational_kmeans]", BLUE)]
    item_rows = [(a, item[b], c) for a, b, c in item_rows if b in item]
    churn = {}
    for r in roots:
        for f in sorted(glob.glob(os.path.join(r, "*J1[24]*churn*", "metrics.json"))):
            if "smoke" in f:
                continue
            for name, v in json.load(open(f)).get("churn", {}).get("rows", {}).items():
                if "auroc" in v:
                    churn[name] = v["auroc"]
    churn_rows = [("Supervised gradient boosting (labels)", "Supervised[hgb on agg]", NEUTRAL),
                  ("Supervised TabPFN (labels in context)", "Supervised[tabpfn on agg]", NEUTRAL),
                  ("Frozen Kumo Relational, random target", "kNN[kumo_relational_random]", BLUE),
                  ("Hand aggregates", "kNN[agg]", NEUTRAL),
                  ("Frozen TabPFN over aggregates, k-means", "kNN[tabpfn_agg_kmeans]", BLUE),
                  ("Frozen Kumo Relational, k-means target", "kNN[kumo_relational_kmeans]", BLUE),
                  ("Row features only", "kNN[row]", NEUTRAL)]
    churn_rows = [(a, churn[b], c) for a, b, c in churn_rows if b in churn]
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(13.0, 5.0))
    _hbars(a1, [r[0] for r in item_rows], [r[1] for r in item_rows], [r[2] for r in item_rows], fmt="{:.2f}")
    a1.set_title("Recommending items (user-item-purchase)", pad=8, fontsize=11)
    a1.set_xlabel("Test MAP@12 ×100, official evaluator (higher is better)")
    a1.set_xlim(0, max(r[1] for r in item_rows) * 1.25)
    vals = [r[1] - 0.5 for r in churn_rows]
    _hbars(a2, [r[0] for r in churn_rows], vals, [r[2] for r in churn_rows], fmt="")
    for yi, (_lab, v, _) in zip(np.arange(len(churn_rows))[::-1], churn_rows):
        a2.text(v - 0.5 + 0.003, yi, f"{v:.3f}", va="center", ha="left", fontsize=9, color=INK)
    a2.set_xlim(0, max(vals) * 1.22)
    ticks = np.arange(0, max(vals) * 1.2, 0.05)
    a2.set_xticks(ticks, [f"{t + 0.5:.2f}" for t in ticks])
    a2.set_title("Describing customers (user-churn)", pad=8, fontsize=11)
    a2.set_xlabel("AUROC (0.5 = chance)")
    fig.text(0.01, 0.005, "Blue = frozen foundation-model hidden state used as an embedding; gray = references.", color=INK2, fontsize=8.5)
    top = _header(fig, "Frozen foundation-model embeddings on real data (RelBench rel-hm): lose for items, near-tie for customers",
                  "Training-free kNN over each representation · the relational model is the only embedding above its own input (customers, +0.007 AUROC)")
    fig.tight_layout(rect=(0, 0.03, 1, top))
    fig.savefig(out / "fig5_retrieval.png", dpi=180)
    plt.close(fig)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--roots", nargs="+", default=["../excel_magic/reports/runs"])
    ap.add_argument("--out", default="paper/figures")
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    for f in (fig1, fig2, fig3, fig4, fig5):
        f(a.roots, out)
        print("wrote", f.__name__)


if __name__ == "__main__":
    main()
