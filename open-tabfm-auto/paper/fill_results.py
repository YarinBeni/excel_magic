"""Regenerate paper/results_auto.md from run directories (local runs/ and/or the cluster's reports/runs/).

  python paper/fill_results.py --roots runs ../excel_magic/reports/runs --out paper/results_auto.md

Tables: T1 backbones (exp01 / J1_backbones), T2 pipeline search by harness x LLM x dataset (search runs),
T3 TabArena vs the paper (comparison_to_paper.csv), T4 backbone transfer, T5 retrieval benchmark (fer).
Every number is read from metrics.json / results.csv; nothing is typed by hand.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import pandas as pd


def load_runs(roots):
    runs = []
    for root in roots:
        for d in sorted(Path(root).glob("*")):
            m, c = d / "metrics.json", d / "config.json"
            if m.exists() and c.exists():
                try:
                    runs.append((d, json.load(open(c)).get("config", {}), json.load(open(m))))
                except Exception:
                    pass
    return runs


def t1_backbones(runs):
    rows = []
    for d, _cfg, met in runs:
        for r in met.get("table", []) if "datasets" in met and "models" in met else []:
            if r.get("score") is None:
                continue
            rows.append({"run": d.name, "dataset": r["dataset"], "model": r["model"].split(":")[0], "metric": r["metric"],
                         "score": r["score"], "cv_s": r.get("cv_seconds")})
    if not rows:
        return "_no backbone-comparison runs yet_"
    df = pd.DataFrame(rows).drop_duplicates(subset=["dataset", "model"], keep="last")
    piv = df.pivot_table(index=["dataset", "metric"], columns="model", values="score")
    return piv.to_markdown(floatfmt=".4f")


def t2_search(runs):
    rows = []
    for d, _cfg, met in runs:
        if "p0_test" not in met or met.get("dataset") is None:
            continue
        agent = met.get("agent", {}) or {}
        rows.append({"dataset": met["dataset"], "backbone": str(met.get("model", "")).split(":")[0], "harness": met.get("harness"),
                     "llm": met.get("llm_model") or agent.get("model") or ("-" if met.get("harness") in ("heuristic", "none") else "?"),
                     "metric": met.get("metric"), "P0 test": met.get("p0_test"), "P* test": met.get("best_test"),
                     "gain %": met.get("test_improvement_pct"), "evals": met.get("n_evals"),
                     "HistGB": met.get("baseline_hgb_test"), "run": d.name})
    if not rows:
        return "_no search runs yet_"
    df = pd.DataFrame(rows).drop_duplicates(subset=["run"]).sort_values(["dataset", "backbone", "harness", "llm"])
    return df.drop(columns=["run"]).to_markdown(index=False, floatfmt=".4f")


def t3_tabarena(roots):
    frames = []
    for root in roots:
        for f in Path(root).glob("*/comparison_to_paper.csv"):
            df = pd.read_csv(f)
            df["run"] = f.parent.name
            frames.append(df)
    if not frames:
        return "_no TabArena-protocol runs yet_"
    df = pd.concat(frames).drop_duplicates(subset=["dataset", "run"], keep="last")
    cols = [c for c in ["dataset", "P0", "P*", "our_gain_%", "tabfm", "tabfm_auto_opus5", "paper_gain_opus5_%", "run"] if c in df.columns]
    return df[cols].to_markdown(index=False, floatfmt=".4f")


def t4_transfer(runs):
    rows = []
    for _d, _cfg, met in runs:
        for r in met.get("table", []) if "models" in met and "n_rows" in met else []:
            if r.get("score") is not None and "pipeline" in r:
                rows.append({"dataset": r["dataset"], "model": r["model"], "pipeline": r["pipeline"], "score": r["score"]})
    if not rows:
        return "_no transfer runs yet_"
    df = pd.DataFrame(rows).drop_duplicates(subset=["dataset", "model", "pipeline"], keep="last")
    piv = df.pivot_table(index=["dataset", "model"], columns="pipeline", values="score")
    if {"P0", "P*"} <= set(piv.columns):
        piv["gain %"] = 100 * (piv["P0"] - piv["P*"]) / piv["P0"]
    return piv.to_markdown(floatfmt=".4f")


def t5_retrieval(roots):
    rows = []
    for root in roots:
        for d in Path(root).glob("*"):
            src = d / "rescored.json" if (d / "rescored.json").exists() else d / "metrics.json"
            if not src.exists():
                continue
            try:
                m = json.load(open(src))
            except Exception:
                continue
            res = m.get("results")
            if isinstance(res, dict):
                items = [(k, v) for k, v in res.items()]
            elif isinstance(res, list):
                items = [(r.get("embedder"), r) for r in res if r.get("status") == "ok"]
            else:
                continue
            for name, r in items:
                if not isinstance(r, dict) or "future_purchase" not in r:
                    continue
                rows.append({"db": m.get("db", "?"), "embedder": name, "seg P@10": r.get("segment", {}).get("P@10"),
                             "seg kNN acc": r.get("segment", {}).get("kNN_acc"), "fut MAP@10": r["future_purchase"]["MAP@10"]})
    if not rows:
        return "_no retrieval runs yet_"
    df = pd.DataFrame(rows).groupby(["db", "embedder"]).agg(["mean", "std", "count"])
    df.columns = [f"{a} {b}" for a, b in df.columns]
    return df.to_markdown(floatfmt=".3f")


def t6_relbench(roots):
    rows = []
    for root in roots:
        files = list(Path(root).glob("*/relbench_rows.json"))
        files += [m for m in Path(root).glob("*/metrics.json") if "splits" in json.load(open(m))]
        for f in files:
            out = json.load(open(f))
            out = out.get("splits", out)
            for split, s in out.items():
                if not isinstance(s, dict) or "rows" not in s or not s.get("official_evaluator", True):
                    continue  # smoke runs on a query subset are not comparable
                for name, r in s["rows"].items():
                    if "map" in r:
                        rows.append({"run": f.parent.name, "split": split, "method": name, "MAP@K x100": 100 * r["map"]})
    if not rows:
        return "_no RelBench runs yet_"
    df = pd.DataFrame(rows).drop_duplicates(subset=["split", "method"], keep="last")
    return df.pivot_table(index="method", columns="split", values="MAP@K x100").to_markdown(floatfmt=".3f")


def t7_churn(roots):
    rows = []
    for root in roots:
        for m in Path(root).glob("*/metrics.json"):
            d = json.load(open(m))
            if "churn" not in d or "smoke" in m.parent.name:
                continue
            for name, r in d["churn"]["rows"].items():
                if "auroc" in r:
                    rows.append({"run": m.parent.name, "method": name, "AUROC": r["auroc"]})
    if not rows:
        return "_no churn-probe runs yet_"
    df = pd.DataFrame(rows).drop_duplicates(subset=["method"], keep="last").sort_values("AUROC", ascending=False)
    return df[["method", "AUROC"]].to_markdown(index=False, floatfmt=".4f")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--roots", nargs="+", default=["runs"])
    ap.add_argument("--out", default="paper/results_auto.md")
    a = ap.parse_args()
    runs = load_runs(a.roots)
    md = ["# Auto-generated results tables", f"_from {len(runs)} runs under {', '.join(a.roots)}_", "",
          "## T1. Frozen backbones with the identity pipeline (3-fold CV, lower is better)", t1_backbones(runs), "",
          "## T2. Pipeline search: harness x LLM x backbone (held-out test error)", t2_search(runs), "",
          "## T3. TabArena protocol vs the paper (per-dataset test error, P0 / P* vs TabFM / TabFM-Auto Opus 5)", t3_tabarena(a.roots), "",
          "## T4. Backbone transfer of discovered pipelines", t4_transfer(runs), "",
          "## T5. Retrieval benchmark (synthetic shop DB; mean/std over seeds)", t5_retrieval(a.roots), "",
          "## T6. RelBench rel-hm user-item-purchase (official evaluator)", t6_relbench(a.roots), "",
          "## T7. RelBench rel-hm user-churn, entity-level kNN probe (AUROC, random customer folds)", t7_churn(a.roots), ""]
    Path(a.out).write_text("\n".join(md))
    print("\n".join(md))


if __name__ == "__main__":
    main()
