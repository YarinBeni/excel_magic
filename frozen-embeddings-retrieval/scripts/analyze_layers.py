"""Combine the three layer studies into one table set and four figures.

  J18  Kumo Relational (graph + in-context transformer), every layer, contexts: random labels / real labels
  J19  TabPFN v2 over flattened relational features, every block, contexts: zeros / random / k-means / real labels
  J21  six tabular FMs (TabPFN v2, TabPFN-2.5, TabICLv2, Kumo Tabular S/M/L), every layer, contexts: random / k-means /
       real labels, plus label-free geometry per layer and CKA within and across models

  python scripts/analyze_layers.py --roots ../excel_magic/reports/runs --out docs/layers

Every number comes from layers_rows.json (latest run per study x task; smoke runs ignored). Test numbers are RelBench's
official evaluator (AUROC). "Val-picked" = the layer with the best validation AUROC; "oracle" = best test AUROC over
layers (an upper bound, not a method).
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

SURFACE, INK, INK2, MUTED, GRID = "#fcfcfb", "#0b0b0b", "#52514e", "#8a8984", "#e6e5e1"
BLUE, ORANGE, AQUA = "#2a78d6", "#eb6834", "#1baf7a"
plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "font.family": "DejaVu Sans", "font.size": 9.5, "text.color": INK, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.edgecolor": GRID, "axes.titleweight": "bold",
    "axes.titlesize": 10.5, "axes.titlelocation": "left", "axes.spines.top": False, "axes.spines.right": False,
})
MODEL_ORDER = ["tabpfn", "tabpfn-2.5", "tabiclv2", "kumo-s", "kumo-m", "kumo-l", "kumo-relational"]
MODEL_LABEL = {"tabpfn": "TabPFN v2", "tabpfn-2.5": "TabPFN-2.5", "tabiclv2": "TabICLv2", "kumo-s": "Kumo Tabular S",
               "kumo-m": "Kumo Tabular M", "kumo-l": "Kumo Tabular L", "kumo-relational": "Kumo Relational"}
TARGET_LABEL = {"random": "random labels", "kmeans": "k-means clusters", "label": "real labels", "zeros": "all zeros"}
TARGET_COLOR = {"random": BLUE, "kmeans": AQUA, "label": ORANGE, "zeros": MUTED}
GEO = ["effective_rank", "intrinsic_dim", "anisotropy", "kmeans_silhouette"]
GEO_LABEL = {"effective_rank": "effective rank", "intrinsic_dim": "intrinsic dimension (TwoNN)",
             "anisotropy": "anisotropy (mean cosine)", "kmeans_silhouette": "k-means silhouette"}


def _latest(roots, pattern, exclude=("smoke",)):
    """layers_rows.json of the latest run per (run name minus timestamp, context/probe split protocol)."""
    best: dict[tuple[str, str], tuple[str, dict]] = {}
    for r in roots:
        for f in glob.glob(os.path.join(r, pattern, "layers_rows.json")):
            run = Path(f).parent.name
            if any(x in run for x in exclude) or not os.path.getsize(f):
                continue
            d = json.load(open(f))
            key = (run.split("_", 1)[1], d.get("ctx_split", "random"))
            if key not in best or best[key][0] < run:
                best[key] = (run, d)
    return [d for _, (_, d) in sorted(best.items())]


def _t(d):  # test AUROC in either schema
    return d["test_auroc"] if "test_auroc" in d else d["test"].get("roc_auc", np.nan)


def load(roots) -> tuple[pd.DataFrame, pd.DataFrame, list[dict]]:
    """Long table: one row per (study, model, task, target, layer); plus one row per (.., target) of summary values."""
    rows, summ, j21 = [], [], []

    def add(study, model, ds, task, target, block, raw=None, split="random"):
        layers = [L for L in block["layers"] if L != "raw_features"]
        for i, L in enumerate(layers):
            v = block["layers"][L]
            rows.append({"split": split, "study": study, "model": model, "task": f"{ds}/{task}", "target": target, "layer": L,
                         "depth": i / max(1, len(layers) - 1), "linear_val": v["linear"]["val_auroc"],
                         "linear_test": _t(v["linear"]), "knn_val": v["knn"]["val_auroc"], "knn_test": _t(v["knn"]),
                         **{g: v.get("geometry", {}).get(g, np.nan) for g in GEO}})
        raw_v = block["layers"].get("raw_features") or raw
        own = block.get("icl_head")
        summ.append({"split": split, "study": study, "model": model, "task": f"{ds}/{task}", "target": target,
                     "raw_linear": _t(raw_v["linear"]) if raw_v else np.nan,
                     "raw_knn": _t(raw_v["knn"]) if raw_v else np.nan,
                     "own": _t(own) if own else np.nan, "own_val": own["val_auroc"] if own else np.nan,
                     "own_probe_train": own.get("probe_train_auroc", np.nan) if own else np.nan,
                     "n_layers": len(layers)})

    for r in _latest(roots, "*J18_layers_*"):
        for mode, block in r["modes"].items():
            add("J18", "kumo-relational", r["dataset"], r["task"], mode, block, split=r.get("ctx_split", "random"))
    for r in _latest(roots, "*J19_tab_layers_*"):
        for target, block in r["targets"].items():
            add("J19", "tabpfn", r["dataset"], r["task"], target, block, split=r.get("ctx_split", "random"))
    for r in _latest(roots, "*J21_model_layers_*"):
        j21.append(r)
        for model, per in r["models"].items():
            for target, block in per.items():
                if "error" not in block:
                    add("J21", model, r["dataset"], r["task"], target, block, split=r.get("ctx_split", "random"))
    L, S = pd.DataFrame(rows), pd.DataFrame(summ)
    # Kumo Relational (J18) reads the subgraph directly and has no raw-feature probe of its own: borrow the probe on the
    # flattened relational features of the same task (J19/J21 build them from the same subgraph sampler)
    for col in ("raw_linear", "raw_knn"):
        ref = S[S["study"] != "J18"].groupby(["split", "task"])[col].mean()
        m = S["study"].eq("J18") & S[col].isna()
        S.loc[m, col] = [ref.get((sp, t), np.nan) for sp, t in zip(S.loc[m, "split"], S.loc[m, "task"])]
    return L, S, j21


def per_run(L: pd.DataFrame, S: pd.DataFrame, probe="linear") -> pd.DataFrame:
    """One row per (study, model, task, target): val-picked layer, last layer, oracle, own prediction, raw features."""
    out = []
    K = ["split", "study", "model", "task", "target"]
    Si = S.set_index(K)
    for key, g in L.groupby(K, sort=False):
        g = g.reset_index(drop=True)
        vp = g.loc[g[f"{probe}_val"].idxmax()]
        orc = g.loc[g[f"{probe}_test"].idxmax()]
        s = Si.loc[key]
        out.append({**dict(zip(K, key)), "own_val": s["own_val"], "own_probe_train": s["own_probe_train"],
                    "last_knn": g["knn_test"].iloc[-1] if probe == "linear" else np.nan,
                    "raw": s[f"raw_{probe}"], "own": s["own"], "first": g[f"{probe}_test"].iloc[0],
                    "last": g[f"{probe}_test"].iloc[-1], "val_layer": vp["layer"], "val_depth": vp["depth"],
                    "val_pick": vp[f"{probe}_test"], "oracle_layer": orc["layer"], "oracle_depth": orc["depth"],
                    "oracle": orc[f"{probe}_test"], "mean_layer": g[f"{probe}_test"].mean()})
    return pd.DataFrame(out)


def geometry_selectors(L: pd.DataFrame) -> pd.DataFrame:
    """Can a label-free statistic pick the layer? Regret = oracle test AUROC - picked layer's test AUROC.
    The direction (pick max or min of the statistic) is chosen leave-one-task-out, so no test label is used to pick it."""
    G = L[(L["study"] == "J21")].dropna(subset=["effective_rank"])
    if G.empty:
        return pd.DataFrame()
    runs = list(G.groupby(["model", "task", "target"], sort=False))
    rec = []
    for (model, task, target), g in runs:
        g = g.reset_index(drop=True)
        orc = g["linear_test"].max()
        r = {"model": model, "task": task, "target": target, "oracle": orc,
             "last": orc - g["linear_test"].iloc[-1], "val": orc - g.loc[g["linear_val"].idxmax(), "linear_test"],
             "random layer": orc - g["linear_test"].mean()}
        for s in GEO:
            x = g[s].to_numpy(dtype=float)
            if np.isfinite(x).sum() < 2:
                continue
            r[f"{s}:max"] = orc - g.loc[int(np.nanargmax(x)), "linear_test"]
            r[f"{s}:min"] = orc - g.loc[int(np.nanargmin(x)), "linear_test"]
            ok = np.isfinite(x) & np.isfinite(g["linear_test"].to_numpy())
            r[f"{s}:rho"] = pd.Series(x[ok]).corr(g["linear_test"][ok], method="spearman") if ok.sum() > 2 else np.nan
        rec.append(r)
    R = pd.DataFrame(rec)
    tasks = R["task"].unique()
    for s in GEO:
        if f"{s}:max" not in R:
            continue
        loto = np.full(len(R), np.nan)
        for t in tasks:
            other = R[R["task"] != t]
            d = "max" if (len(other) == 0 or other[f"{s}:max"].mean() <= other[f"{s}:min"].mean()) else "min"
            m = (R["task"] == t).to_numpy()
            loto[m] = R.loc[m, f"{s}:{d}"]
        R[f"{s}:loto"] = loto
    return R


def fmt(x, nd=3):
    return "" if x is None or (isinstance(x, float) and not np.isfinite(x)) else f"{x:.{nd}f}"


def protocol_table(P: pd.DataFrame) -> str:
    """Same (study, model, task, context) under both context/probe splits: does the random split penalise late layers?"""
    K = ["study", "model", "task", "target"]
    a, b = P[P["split"] == "random"].set_index(K), P[P["split"] == "time"].set_index(K)
    both = a.index.intersection(b.index)
    if not len(both):
        return ""
    md = ["", "## Protocol check: random vs time split of context and probe-train rows", "",
          "Random split: probe-train rows share time periods (and entities) with the context. Time split: probe-train rows "
          "are strictly later than every context row, as validation and test rows are. Mean test AUROC over the runs "
          "present under both protocols.", "",
          "| context | runs | split | last layer (linear) | last layer (kNN) | val-picked (linear) | own prediction | own prediction on probe-train rows |",
          "|---|---|---|---|---|---|---|---|"]
    for tg in ["random", "kmeans", "label", "zeros"]:
        idx = [i for i in both if i[3] == tg]
        if not idx:
            continue
        for name, X in (("random", a), ("time", b)):
            x = X.loc[idx]
            md.append(f"| {TARGET_LABEL.get(tg, tg)} | {len(idx)} | {name} | {fmt(x['last'].mean())} | {fmt(x['last_knn'].mean())} | "
                      f"{fmt(x['val_pick'].mean())} | {fmt(x['own'].mean())} | {fmt(x['own_probe_train'].mean())} |")
    return "\n".join(md) + "\n"


def tables(P: pd.DataFrame, R: pd.DataFrame, j21: list[dict], split: str = "time") -> str:
    md = ["# Layer study: which layer of a frozen tabular foundation model gives the best entity embedding?", "",
          f"Protocol: **{split} split** of context and probe-train rows "
          + ("(probe-train rows strictly later than the context, like validation and test rows)." if split == "time"
             else "(random; probe-train rows can share periods and entities with the context: see the protocol check)."), "",
          "Generated by `scripts/analyze_layers.py` from the run directories. Test AUROC from RelBench's official "
          "evaluator. Linear probe (logistic regression) on the frozen layer output; the layer is picked on "
          "validation. *Oracle* = best test layer (upper bound only). Kumo Relational reads the subgraph itself; its "
          "raw-features column is the probe on the flattened relational features of the same task (from J19/J21).", ""]
    tasks = sorted(P["task"].unique())
    md += ["## Coverage", "", "| study | model | tasks | targets |", "|---|---|---|---|"]
    for (st, m), g in P.groupby(["study", "model"], sort=False):
        md.append(f"| {st} | {MODEL_LABEL.get(m, m)} | {g['task'].nunique()} | {', '.join(sorted(g['target'].unique()))} |")
    md += ["", "## Per task: val-picked layer vs last layer vs the model's own prediction (real-label context)", "",
           "| task | study | model | raw features | own prediction | last layer | val-picked layer (test) | oracle layer |",
           "|---|---|---|---|---|---|---|---|"]
    lab = P[P["target"] == "label"]
    for t in tasks:
        for _, r in lab[lab["task"] == t].iterrows():
            md.append(f"| {t} | {r['study']} | {MODEL_LABEL.get(r['model'], r['model'])} | {fmt(r['raw'])} | {fmt(r['own'])} | "
                      f"{fmt(r['last'])} | {r['val_layer']} ({fmt(r['val_pick'])}) | {r['oracle_layer']} ({fmt(r['oracle'])}) |")
    md += ["", "## Averages over tasks (each cell: mean test AUROC; count of tasks in brackets)", "",
           "| study | model | context | n | raw features | own prediction | first layer | last layer | val-picked | oracle | val-picked minus last | val-picked minus own |",
           "|---|---|---|---|---|---|---|---|---|---|---|---|"]
    for (st, m, tg), g in P.groupby(["study", "model", "target"], sort=False):
        md.append(f"| {st} | {MODEL_LABEL.get(m, m)} | {TARGET_LABEL.get(tg, tg)} | {len(g)} | {fmt(g['raw'].mean())} | "
                  f"{fmt(g['own'].mean())} | {fmt(g['first'].mean())} | {fmt(g['last'].mean())} | {fmt(g['val_pick'].mean())} | "
                  f"{fmt(g['oracle'].mean())} | {fmt((g['val_pick'] - g['last']).mean(), 3)} | "
                  f"{fmt((g['val_pick'] - g['own']).mean(), 3)} |")
    # where is the best layer?
    md += ["", "## Where the val-picked layer sits (relative depth 0 = first layer, 1 = last)", "",
           "| study | model | context | median depth | share of runs where an inner layer beats the last on test |", "|---|---|---|---|---|"]
    for (st, m, tg), g in P.groupby(["study", "model", "target"], sort=False):
        md.append(f"| {st} | {MODEL_LABEL.get(m, m)} | {TARGET_LABEL.get(tg, tg)} | {fmt(g['val_depth'].median(), 2)} | "
                  f"{(g['val_pick'] > g['last']).mean():.0%} ({len(g)}) |")
    if not R.empty:
        sel = ["last", "random layer", "val"] + [f"{s}:loto" for s in GEO if f"{s}:loto" in R]
        md += ["", "## Can label-free geometry pick the layer? (J21; regret = oracle minus picked, lower is better)", "",
               "Direction (pick the max or the min of the statistic) chosen leave-one-task-out.", "",
               "| selector | mean regret | median regret | runs |", "|---|---|---|---|"]
        for s in sel:
            name = {"last": "last layer", "random layer": "average layer", "val": "validation labels (supervised)"}.get(s, GEO_LABEL.get(s.split(":")[0], s) + " (label-free)")
            md.append(f"| {name} | {fmt(R[s].mean())} | {fmt(R[s].median())} | {R[s].notna().sum()} |")
        md += ["", "| statistic | mean Spearman with layer test AUROC | share of runs with rho > 0 |", "|---|---|---|"]
        for s in GEO:
            if f"{s}:rho" in R:
                md.append(f"| {GEO_LABEL[s]} | {fmt(R[f'{s}:rho'].mean(), 2)} | {(R[f'{s}:rho'] > 0).mean():.0%} |")
    if j21:
        md += ["", "## Cross-model CKA (J21, real-label context, mean over tasks)", ""]
        mats = {}
        for r in j21:
            c = r.get("cka_models", {}).get("label")
            if c:
                for i, a in enumerate(c["keys"]):
                    for j, b in enumerate(c["keys"]):
                        mats.setdefault((a, b), []).append(c["matrix"][i][j])
        keys = [k for k in dict.fromkeys(a for a, _ in mats)]
        md += ["| | " + " | ".join(keys) + " |", "|---" * (len(keys) + 1) + "|"]
        for a in keys:
            md.append(f"| {a} | " + " | ".join(fmt(np.mean(mats.get((a, b), [np.nan])), 2) for b in keys) + " |")
    return "\n".join(md) + "\n"


def _header(fig, title, sub):
    h = fig.get_figheight()
    fig.text(0.012, 1 - 0.16 / h, title, fontsize=12, fontweight="bold", color=INK, va="top")
    fig.text(0.012, 1 - 0.42 / h, sub, fontsize=8.8, color=INK2, va="top")
    return 1 - 0.68 / h


def fig_depth(L: pd.DataFrame, S: pd.DataFrame, out: Path):
    """Small multiples: one panel per model; x = relative depth, y = test AUROC minus raw features; one line per context."""
    studies = ["J18", "J21"] if (L["study"].eq("J21") & L["model"].eq("tabpfn")).any() else ["J18", "J19", "J21"]
    D = L[L["study"].isin(studies)].merge(S, on=["split", "study", "model", "task", "target"])
    if D.empty:
        return
    D["gain"] = D["linear_test"] - D["raw_linear"]
    models = [m for m in MODEL_ORDER if m in D["model"].unique()]
    nc = 4 if len(models) > 4 else len(models)
    nr = int(np.ceil(len(models) / nc))
    fig, axs = plt.subplots(nr, nc, figsize=(max(8.4, 2.55 * nc + 0.4), 2.15 * nr + 1.2), sharey=True, squeeze=False)
    top = _header(fig, "Which layer holds the signal? Linear probe on each layer, minus raw features",
                  "Test AUROC gain over a probe on the raw relational features (0 = no better). Mean over RelBench tasks. "
                  "x: relative depth.")
    for ax, m in zip(axs.flat, models):
        g = D[D["model"] == m]
        for tg in ["zeros", "random", "kmeans", "label"]:
            h = g[g["target"] == tg]
            if h.empty:
                continue
            h = h.assign(db=(h["depth"] * 12).round() / 12).groupby("db")["gain"].mean()
            ax.plot(h.index, h.values, color=TARGET_COLOR[tg], lw=2, marker="o", ms=3, label=TARGET_LABEL[tg])
        own = S[(S["model"] == m) & (S["target"] == "label")]
        if own["own"].notna().any():
            ax.axhline((own["own"] - own["raw_linear"]).mean(), color=ORANGE, lw=1.2, ls="--")
        ax.axhline(0, color=MUTED, lw=0.8)
        ax.set_title(f"{MODEL_LABEL[m]}  (n={g['task'].nunique()})", fontsize=9.5)
        ax.grid(axis="y", color=GRID, lw=0.6)
        ax.set_xticks([0, 0.5, 1])
    for ax in axs.flat[len(models):]:
        ax.axis("off")
    for ax in axs[:, 0]:
        ax.set_ylabel("AUROC minus raw")
    hs, ls_ = [], []
    for ax in axs.flat[:len(models)]:
        for h, l in zip(*ax.get_legend_handles_labels()):
            if l not in ls_:
                hs.append(h); ls_.append(l)
    hs.append(plt.Line2D([], [], color=ORANGE, ls="--", lw=1.2)); ls_.append("own prediction (real labels)")
    bottom = 0
    if len(models) < nr * nc:
        axs.flat[len(models)].legend(hs, ls_, loc="center", frameon=False, fontsize=8.5, title="context labels",
                                     title_fontsize=8.5)
    else:
        fig.legend(hs, ls_, loc="lower center", frameon=False, fontsize=8.5, ncol=len(hs))
        bottom = 0.35 / fig.get_figheight()
    fig.tight_layout(rect=(0, bottom, 1, top))
    fig.savefig(out / "fig_layers_depth.png", dpi=170)
    plt.close(fig)


def fig_geometry(L: pd.DataFrame, out: Path):
    """Heatmaps: rows = models, columns = relative depth bins, colour = statistic (real-label context, mean over tasks)."""
    G = L[(L["study"] == "J21") & (L["target"] == "label")].dropna(subset=["effective_rank"])
    if G.empty:
        return
    G = G.assign(db=(G["depth"] * 6).round().astype(int))
    models = [m for m in MODEL_ORDER if m in G["model"].unique()]
    fig, axs = plt.subplots(1, 4, figsize=(11.2, 0.42 * len(models) + 1.9))
    top = _header(fig, "Label-free geometry of each layer (real-label context)",
                  "Mean over RelBench tasks. Columns: relative depth, first layer (0) to last (1). Darker = larger within that model's row.")
    for ax, s in zip(axs, GEO):
        M = G.pivot_table(index="model", columns="db", values=s, aggfunc="mean").reindex(models)
        A = M.to_numpy(dtype=float)
        lo, hi = np.nanmin(A, axis=1, keepdims=True), np.nanmax(A, axis=1, keepdims=True)
        Z = (A - lo) / (hi - lo + 1e-12)  # scale differs by model width: colour within each row
        ax.imshow(Z, aspect="auto", cmap="Blues", vmin=0, vmax=1)
        for i in range(M.shape[0]):
            for j in range(M.shape[1]):
                v = A[i, j]
                if np.isfinite(v):
                    ax.text(j, i, f"{v:.2g}" if abs(v) < 100 else f"{v:.0f}", ha="center", va="center", fontsize=6.5,
                            color="white" if Z[i, j] > 0.6 else INK)
        ax.set_xticks(range(M.shape[1]), [f"{c / 6:.1f}" for c in M.columns], fontsize=7)
        ax.set_yticks(range(len(models)), [MODEL_LABEL[m] for m in models] if ax is axs[0] else [""] * len(models), fontsize=8)
        ax.set_title(GEO_LABEL[s], fontsize=9)
        for sp in ax.spines.values():
            sp.set_visible(False)
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig_layers_geometry.png", dpi=170)
    plt.close(fig)


def fig_selectors(R: pd.DataFrame, out: Path):
    if R.empty:
        return
    sel = {"val": "validation labels", "last": "last layer", "random layer": "average layer"}
    sel.update({f"{s}:loto": GEO_LABEL[s] for s in GEO if f"{s}:loto" in R})
    v = pd.Series({lab: R[k].mean() for k, lab in sel.items()}).sort_values()
    fig, ax = plt.subplots(figsize=(7.2, 0.36 * len(v) + 1.3))
    top = _header(fig, "Picking the layer: label-free statistics vs validation labels",
                  "Mean test-AUROC regret vs the best layer (lower is better). Geometry direction chosen leave-one-task-out.")
    ax.barh(range(len(v)), v.values, color=[ORANGE if k == "validation labels" else BLUE for k in v.index], height=0.62)
    for i, x in enumerate(v.values):
        ax.text(x + v.max() * 0.01, i, f"{x:.3f}", va="center", fontsize=8, color=INK2)
    ax.set_yticks(range(len(v)), v.index)
    ax.invert_yaxis()
    ax.set_xlabel("mean regret (AUROC)")
    ax.grid(axis="x", color=GRID, lw=0.6)
    ax.set_xlim(0, v.max() * 1.15)
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig_layers_selectors.png", dpi=170)
    plt.close(fig)


def fig_cka(j21: list[dict], out: Path):
    """Within-model layer-by-layer CKA (one task, real labels) for each model."""
    if not j21:
        return
    r = j21[0]
    models = [m for m in MODEL_ORDER if m in r["models"] and "cka_layers" in r["models"][m].get("label", {})]
    if not models:
        return
    nc = min(3, len(models)); nr = int(np.ceil(len(models) / nc))
    fig, axs = plt.subplots(nr, nc, figsize=(3.2 * nc, 2.9 * nr + 0.7), squeeze=False)
    top = _header(fig, "Layer-to-layer similarity inside each model",
                  f"Linear CKA, {r['dataset']}/{r['task']}, real-label context. Dark = similar (1), light = unrelated (0).")
    for ax, m in zip(axs.flat, models):
        c = r["models"][m]["label"]["cka_layers"]
        ax.imshow(np.array(c["matrix"], dtype=float), vmin=0, vmax=1, cmap="Blues")
        n = len(c["layers"])
        ticks = [0, n // 2, n - 1]
        ax.set_xticks(ticks, [c["layers"][i] for i in ticks], fontsize=6.5)
        ax.set_yticks(ticks, [c["layers"][i] for i in ticks], fontsize=6.5)
        ax.set_title(MODEL_LABEL[m], fontsize=9)
        for sp in ax.spines.values():
            sp.set_visible(False)
    for ax in axs.flat[len(models):]:
        ax.axis("off")
    fig.tight_layout(rect=(0, 0, 1, top))
    fig.savefig(out / "fig_layers_cka.png", dpi=170)
    plt.close(fig)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--roots", nargs="+", default=["../excel_magic/reports/runs"])
    ap.add_argument("--out", default="docs/layers")
    ap.add_argument("--ctx-split", default="time", choices=["time", "random"],
                    help="protocol for the main tables and figures (falls back to random if no time-split runs exist)")
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    L_all, S_all, j21_all = load(a.roots)
    if L_all.empty:
        raise SystemExit("no layer runs found")
    P_all = per_run(L_all, S_all)
    split = a.ctx_split if L_all["split"].eq(a.ctx_split).any() else "random"
    L, S = L_all[L_all["split"] == split], S_all[S_all["split"] == split]
    P = P_all[P_all["split"] == split]
    j21 = [r for r in j21_all if r.get("ctx_split", "random") == split]
    R = geometry_selectors(L)
    L_all.to_csv(out / "layers_long.csv", index=False)
    P_all.to_csv(out / "layers_per_run.csv", index=False)
    if not R.empty:
        R.to_csv(out / "layers_selectors.csv", index=False)
    (out / "LAYERS.md").write_text(tables(P, R, j21, split) + protocol_table(P_all))
    fig_depth(L, S, out)
    fig_geometry(L, out)
    fig_selectors(R, out)
    fig_cka(j21, out)
    print((out / "LAYERS.md").read_text())


if __name__ == "__main__":
    main()
