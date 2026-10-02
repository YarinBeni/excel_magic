"""Exp01: how good / how slow is the frozen TFM on CPU vs classical baselines, identity pipeline P0.

Example: python examples/01_tfm_vs_baselines.py --datasets wine,breast_cancer,diabetes,synth_physics,synth_entities \
             --models tabpfn:n_estimators=4,hgb,rf,logreg
"""
from __future__ import annotations

import argparse
import tempfile
import time
import warnings
from pathlib import Path

from tabfm_auto.data import load_task
from tabfm_auto.harness import evaluate_cv
from tabfm_auto.logging_utils import RunLogger
from tabfm_auto.pipeline import IDENTITY_PIPELINE

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--name", default="exp01_tfm_baselines")
    ap.add_argument("--datasets", default="wine,breast_cancer,diabetes,synth_physics,synth_entities")
    ap.add_argument("--models", default="tabpfn:n_estimators=4,hgb,rf,logreg,dummy")
    ap.add_argument("--n-folds", type=int, default=3)
    ap.add_argument("--seed", type=int, default=0)
    a = ap.parse_args()
    cfg = vars(a)
    with RunLogger(a.name, cfg) as run:
        p0 = Path(tempfile.mkdtemp()) / "pipeline.py"
        p0.write_text(IDENTITY_PIPELINE)
        rows = []
        for ds in a.datasets.split(","):
            task = load_task(ds)
            run.event("dataset", **task.summary())
            for spec in a.models.split(","):
                t0 = time.time()
                r = evaluate_cv(p0, task.X, task.y, task.task_type, spec, n_folds=a.n_folds, seed=a.seed, log=run)
                row = {"dataset": ds, "task_type": task.task_type, "n_rows": len(task.X), "model": spec,
                       "metric": r["metric"], "score": r.get("score"), "score_std": r.get("score_std"),
                       "status": r["status"], "error": r.get("error"), "secondary": r.get("secondary"),
                       "cv_seconds": round(time.time() - t0, 2)}
                rows.append(row)
                run.event("result", **row)
                run.info("%-16s %-24s %s=%s (%.1fs) %s", ds, spec, r["metric"], None if r.get("score") is None
                         else round(r["score"], 5), row["cv_seconds"], r.get("error", ""))
        run.save_json("results.json", rows)
        # markdown table
        lines = ["| dataset | model | metric | score | std | cv s |", "|---|---|---|---|---|---|"]
        for r in rows:
            s = "ERR" if r["score"] is None else f"{r['score']:.4f}"
            lines.append(f"| {r['dataset']} | {r['model']} | {r['metric']} | {s} | "
                         f"{(r['score_std'] or 0):.4f} | {r['cv_seconds']} |")
        run.save_text("results.md", "\n".join(lines))
        run.finish({"n_results": len(rows), "datasets": a.datasets, "models": a.models,
                    "table": rows})
        print("\n".join(lines))


if __name__ == "__main__":
    main()
