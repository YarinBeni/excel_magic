"""Candidate selection rules for the end of a search.

The judge is k-fold CV on a few hundred rows; the best-CV candidate of a search is often *not* better than P0 on other
splits of the same data (open-tabfm-auto RESULTS: ~50% of the official splits improve, a coin flip). These rules make
the final pick conservative instead of greedy, using the per-fold scores every evaluation already records:

  best      the CV-best candidate (the paper's rule)
  gatedZ    the CV-best candidate among those whose paired per-fold improvement over P0 is > Z standard errors
            (Z = 1, 2); otherwise P0. A candidate must have the same folds as P0 (same seed), which the harness guarantees.
  ensK      the K best candidates by CV (P0 included if it ranks), averaged at prediction time.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from ..pipeline import load_pipeline, run_pipeline
from .metrics import primary_metric_name, score

RULES = ("p0", "best", "gated1", "gated2", "ens3")


def _ok(evals: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [r for r in evals if r.get("status") == "ok" and r.get("score") is not None]


def paired_stats(p0: dict[str, Any], cand: dict[str, Any]) -> tuple[float, float]:
    """(mean, stderr) of the per-fold improvement P0 - candidate (positive = candidate better)."""
    a, b = np.asarray(p0.get("folds") or [], float), np.asarray(cand.get("folds") or [], float)
    if len(a) == 0 or len(a) != len(b):
        return float(p0["score"] - cand["score"]), float("inf")
    d = a - b
    se = float(d.std(ddof=1) / np.sqrt(len(d))) if len(d) > 1 else float("inf")
    return float(d.mean()), se


def select(evals: list[dict[str, Any]], rule: str = "best") -> list[str]:
    """Candidate file names (``eval_NNN.py``) chosen by ``rule``; one name except for ensK."""
    ok = _ok(evals)
    if not ok:
        return []
    p0 = ok[0]
    ranked = sorted(ok, key=lambda r: r["score"])
    if rule == "p0":
        return [p0["candidate"]]
    if rule == "best":
        return [ranked[0]["candidate"]]
    if rule.startswith("gated"):
        z = float(rule[5:] or 1)
        passing = [r for r in ranked if r is not p0 and (lambda m, se: se != float("inf") and m > z * se)(*paired_stats(p0, r))]
        return [passing[0]["candidate"]] if passing else [p0["candidate"]]
    if rule.startswith("ens"):
        k = int(rule[3:] or 3)
        return [r["candidate"] for r in ranked[:k]]
    raise ValueError(f"unknown selection rule {rule!r}; use one of {RULES}")


def evaluate_holdout_ensemble(pipeline_paths: list[str | Path], X_train: pd.DataFrame, y_train: pd.Series,
                              X_test: pd.DataFrame, y_test: pd.Series, task_type: str, model_spec: str = "tabpfn",
                              seed: int = 0, max_rows: int = 10000, log: Any = None) -> dict[str, Any]:
    """Average the predictions of several pipelines (each fit on the full training split) and score once."""
    metric = primary_metric_name(task_type)
    res: dict[str, Any] = {"metric": metric, "status": "ok", "n_members": len(pipeline_paths)}
    try:
        classes = np.unique(y_train) if task_type != "regression" else None
        preds = []
        for p in pipeline_paths:
            pred, _ = run_pipeline(load_pipeline(p), X_train, y_train, X_test, task_type, model_spec, max_rows=max_rows,
                                   seed=seed, log=log)
            preds.append(pred)
        if task_type == "regression":
            avg = pd.concat([pd.Series(np.asarray(p, float)) for p in preds], axis=1).mean(axis=1)
        else:
            frames = [pd.DataFrame(p) if not isinstance(p, pd.DataFrame) else p for p in preds]
            cols = frames[0].columns
            avg = sum(f.reindex(columns=cols).fillna(0).to_numpy(float) for f in frames) / len(frames)
            avg = pd.DataFrame(avg, columns=cols)
        m = score(task_type, y_test, avg, classes)
        res.update({"score": m[metric], "secondary": m})
    except Exception as e:
        res.update({"status": "error", "score": None, "error": f"{type(e).__name__}: {e}"})
    return res
