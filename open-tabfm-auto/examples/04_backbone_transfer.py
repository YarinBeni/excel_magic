"""Exp05: backbone transfer (paper Section B.3 / Figure 5).

Takes the pipelines discovered by earlier search runs (``runs/*/best_pipeline.py`` + their ``candidates/eval_001.py``
identity pipeline and ``split.json``) and re-scores P0 and P* on the SAME held-out split with every backbone that is
runnable right now. Nothing is searched again, so this is cheap, and it is the first thing to run after new weights
arrive: ``tabfm-models download kumo-tabular-s tabicl && python examples/04_backbone_transfer.py``.

Example: python examples/04_backbone_transfer.py --models tabpfn,kumo-tabular-s,tabicl,hgb
"""
from __future__ import annotations

import argparse
import json
import warnings
from pathlib import Path

import numpy as np

from tabfm_auto.data import load_task
from tabfm_auto.harness.evaluator import evaluate_holdout
from tabfm_auto.logging_utils import RUNS_ROOT, RunLogger
from tabfm_auto.models import is_available, split_model_specs

warnings.filterwarnings("ignore")


def find_search_runs(root: Path) -> list[Path]:
    return sorted(d for d in root.glob("*search_*") if (d / "best_pipeline.py").exists() and (d / "split.json").exists())


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--runs", default=None, help="comma-separated run dirs (default: every search run with a best_pipeline.py)")
    ap.add_argument("--models", default="tabpfn,tabpfn-2.5,tabpfn-3,tabicl,kumo-tabular-s,exaone,hgb")
    ap.add_argument("--n-estimators", type=int, default=4)
    ap.add_argument("--skip-unavailable", action="store_true", default=True)
    ap.add_argument("--name", default="exp05_backbone_transfer")
    a = ap.parse_args()
    runs = [Path(r) for r in a.runs.split(",")] if a.runs else find_search_runs(RUNS_ROOT)
    with RunLogger(a.name, vars(a) | {"runs": [str(r) for r in runs]}) as run:
        models = []
        for m in split_model_specs(a.models):
            if is_available(m):
                models.append(m)
            else:
                run.warning("skipping %s (not runnable here; see `tabfm-models status`)", m)
        rows = []
        for rd in runs:
            cfg = json.loads((rd / "config.json").read_text())["config"]
            task = load_task(cfg["dataset"])
            split = json.loads((rd / "split.json").read_text())
            tr, te = np.asarray(split["train_idx"]), np.asarray(split["test_idx"])
            Xtr, ytr = task.X.iloc[tr].reset_index(drop=True), task.y.iloc[tr].reset_index(drop=True)
            Xte, yte = task.X.iloc[te].reset_index(drop=True), task.y.iloc[te].reset_index(drop=True)
            for m in models:
                if m in ("hgb", "rf", "logreg", "lightgbm", "dummy") or ":" in m:
                    spec = m
                else:
                    spec = f"{m}:n_estimators={a.n_estimators}"
                for label, pipe in (("P0", rd / "workspace" / "candidates" / "eval_001.py"), ("P*", rd / "best_pipeline.py")):
                    r = evaluate_holdout(pipe, Xtr, ytr, Xte, yte, task.task_type, spec, cfg.get("seed", 0),
                                         cfg.get("max_rows", 10000), run)
                    row = {"search_run": rd.name, "dataset": task.name, "model": m, "pipeline": label,
                           "metric": r["metric"], "score": r.get("score"), "status": r["status"], "error": r.get("error"),
                           "seconds": r["elapsed_s"]}
                    rows.append(row)
                    run.event("result", **row)
                    run.info("%-14s %-16s %-3s %s=%s (%.1fs) %s", task.name, m, label, r["metric"],
                             None if r.get("score") is None else round(r["score"], 5), r["elapsed_s"], r.get("error", ""))
        # table: one line per (dataset, model) with P0, P* and the gain
        lines = ["| dataset | model | metric | P0 test | P* test | gain % |", "|---|---|---|---|---|---|"]
        by = {}
        for r in rows:
            by.setdefault((r["dataset"], r["model"], r["metric"]), {})[r["pipeline"]] = r["score"]
        for (ds, m, met), d in by.items():
            p0, ps = d.get("P0"), d.get("P*")
            gain = "" if p0 is None or ps is None or p0 == 0 else f"{100 * (p0 - ps) / p0:+.1f}"
            lines.append(f"| {ds} | {m} | {met} | {'ERR' if p0 is None else f'{p0:.4f}'} | {'ERR' if ps is None else f'{ps:.4f}'} | {gain} |")
        run.save_text("results.md", "\n".join(lines))
        run.finish({"models": models, "n_rows": len(rows), "table": rows})
        print("\n".join(lines))


if __name__ == "__main__":
    main()
