"""``tabfm-eval``: the only evaluation entry point the pipeline-search agent may call.

Run inside a workspace directory that contains ``task.json``, ``train.parquet`` and ``pipeline.py``.
Scores ``pipeline.py`` with 3-fold CV on the training split, appends a record to ``evals.jsonl`` and
snapshots the pipeline into ``candidates/``. Prints a short human summary.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
from pathlib import Path

import pandas as pd

from .evaluator import evaluate_cv
from .sandbox import run_isolated


def _read_workspace(ws: Path):
    task = json.loads((ws / "task.json").read_text())
    df = pd.read_parquet(ws / "train.parquet")
    y = df[task["target"]]
    X = df.drop(columns=[task["target"]])
    return task, X, y


def _inner(ws: Path, pipeline: Path) -> None:
    task, X, y = _read_workspace(ws)
    res = evaluate_cv(pipeline, X, y, task["task_type"], task.get("model_spec", "tabpfn"),
                      n_folds=task.get("n_folds", 3), seed=task.get("seed", 0), max_rows=task.get("max_rows", 10000))
    print(json.dumps(res))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="tabfm-eval", description=__doc__)
    ap.add_argument("--workspace", default=".", help="directory with task.json / train.parquet / pipeline.py")
    ap.add_argument("--pipeline", default="pipeline.py")
    ap.add_argument("--timeout", type=int, default=None, help="seconds (default: task.json eval_timeout_s or 900)")
    ap.add_argument("--inner", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--json", action="store_true", help="print the full JSON record")
    a = ap.parse_args(argv)
    ws = Path(a.workspace).resolve()
    pipeline = (ws / a.pipeline) if not Path(a.pipeline).is_absolute() else Path(a.pipeline)
    if a.inner:
        _inner(ws, pipeline)
        return 0

    task = json.loads((ws / "task.json").read_text())
    evals_path = ws / "evals.jsonl"
    prior = [json.loads(l) for l in evals_path.read_text().splitlines() if l.strip()] if evals_path.exists() else []
    budget = task.get("budget_evals")
    if budget is not None and len(prior) >= budget:
        print(f"BUDGET EXHAUSTED: {len(prior)}/{budget} evaluations used. Stop editing and finish.")
        return 2
    timeout = a.timeout or task.get("eval_timeout_s", 900)
    n = len(prior) + 1
    src = pipeline.read_text()
    (ws / "candidates").mkdir(exist_ok=True)
    snap = ws / "candidates" / f"eval_{n:03d}.py"
    shutil.copy(pipeline, snap)
    t0 = time.time()
    res = run_isolated(["-m", "tabfm_auto.harness.cli", "--inner", "--workspace", str(ws), "--pipeline", str(pipeline)],
                       cwd=ws, timeout_s=timeout, threads=task.get("threads", 4))
    res.update({"eval_id": n, "candidate": snap.name, "wall_s": round(time.time() - t0, 2),
                "pipeline_sha": __import__("hashlib").sha1(src.encode()).hexdigest()[:10]})
    with open(evals_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(res) + "\n")

    ok = [r for r in prior if r.get("status") == "ok" and r.get("score") is not None]
    p0 = ok[0]["score"] if ok else None
    best_prev = min((r["score"] for r in ok), default=None)
    metric = res.get("metric", task.get("metric", "score"))
    if a.json:
        print(json.dumps(res, indent=2))
    if res.get("status") == "ok":
        s = res["score"]
        flag = "NEW BEST" if best_prev is None or s < best_prev else ""
        print(f"eval #{n}: {metric} = {s:.5f} (+/- {res.get('score_std', 0):.4f} over {res.get('n_folds')} folds) "
              f"{flag}")
        print(f"  features={res.get('n_features')} views={res.get('n_views')} time={res['wall_s']}s | "
              f"P0={p0 if p0 is None else round(p0, 5)} best_before={best_prev if best_prev is None else round(best_prev, 5)}")
        sec = res.get("secondary", {})
        if sec:
            print("  secondary: " + ", ".join(f"{k}={v:.4f}" for k, v in sec.items()))
    else:
        print(f"eval #{n}: FAILED -> {res.get('error')}")
        tb = res.get("traceback") or res.get("stderr") or ""
        if tb:
            print("  " + "\n  ".join(tb.strip().splitlines()[-12:]))
    if budget is not None:
        print(f"  budget: {n}/{budget} evaluations used")
    return 0 if res.get("status") == "ok" else 1


if __name__ == "__main__":
    sys.exit(main())
