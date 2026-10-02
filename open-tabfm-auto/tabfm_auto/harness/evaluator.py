"""3-fold cross-validation of a candidate pipeline on the agent-visible training split, and the
one-shot held-out evaluation the harness runs after search."""
from __future__ import annotations

import time
import traceback
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
from sklearn.model_selection import KFold, StratifiedKFold

from ..pipeline import load_pipeline, run_pipeline
from .metrics import primary_metric_name, score


def _folds(y: pd.Series, task_type: str, n_folds: int, seed: int):
    if task_type == "regression":
        return KFold(n_folds, shuffle=True, random_state=seed).split(y)
    return StratifiedKFold(n_folds, shuffle=True, random_state=seed).split(np.zeros(len(y)), y)


def evaluate_cv(pipeline_path: str | Path, X: pd.DataFrame, y: pd.Series, task_type: str,
                model_spec: str = "tabpfn", n_folds: int = 3, seed: int = 0, max_rows: int = 10000,
                log: Any = None, n_repeats: int = 1) -> dict[str, Any]:
    """Returns {"status": "ok"|"error", "score": float, "metric": str, "folds": [...], ...}.
    ``n_repeats`` > 1 repeats the k-fold split with seeds seed, seed+1, ... and averages all folds: a less noisy
    judge for small tables (every open-LLM search overfit the 3-fold signal on a 569-row table)."""
    t0 = time.time()
    metric = primary_metric_name(task_type)
    n_repeats = max(1, int(n_repeats))
    res: dict[str, Any] = {"metric": metric, "n_folds": n_folds, "n_repeats": n_repeats, "status": "ok"}
    try:
        mod = load_pipeline(pipeline_path)
        fold_scores, fold_info, all_metrics = [], [], []
        classes = np.unique(y) if task_type != "regression" else None
        splits = [(r, tr, va) for r in range(n_repeats) for tr, va in _folds(y, task_type, n_folds, seed + r)]
        for k, (r, tr, va) in enumerate(splits):
            pred, info = run_pipeline(mod, X.iloc[tr], y.iloc[tr], X.iloc[va], task_type, model_spec,
                                      max_rows=max_rows, seed=seed, log=log)
            m = score(task_type, y.iloc[va], pred, classes)
            fold_scores.append(m[metric])
            all_metrics.append(m)
            fold_info.append(info)
            if log is not None:
                log.event("cv_fold", fold=k, repeat=r, metrics=m, info=info)
        res["score"] = float(np.mean(fold_scores))
        res["score_std"] = float(np.std(fold_scores))
        res["folds"] = fold_scores
        res["secondary"] = {k: float(np.mean([m[k] for m in all_metrics])) for k in all_metrics[0]}
        res["n_features"] = fold_info[0].get("n_features")
        res["n_views"] = fold_info[0].get("n_views")
        res["model_spec"] = fold_info[0].get("model_spec")
    except Exception as e:
        res["status"] = "error"
        res["score"] = None
        res["error"] = f"{type(e).__name__}: {e}"
        res["traceback"] = traceback.format_exc()[-3000:]
    res["elapsed_s"] = round(time.time() - t0, 2)
    return res


def evaluate_holdout(pipeline_path: str | Path, X_train: pd.DataFrame, y_train: pd.Series,
                     X_test: pd.DataFrame, y_test: pd.Series, task_type: str, model_spec: str = "tabpfn",
                     seed: int = 0, max_rows: int = 10000, log: Any = None) -> dict[str, Any]:
    t0 = time.time()
    metric = primary_metric_name(task_type)
    res: dict[str, Any] = {"metric": metric, "status": "ok"}
    try:
        mod = load_pipeline(pipeline_path)
        classes = np.unique(y_train) if task_type != "regression" else None
        pred, info = run_pipeline(mod, X_train, y_train, X_test, task_type, model_spec, max_rows=max_rows,
                                  seed=seed, log=log)
        m = score(task_type, y_test, pred, classes)
        res.update({"score": m[metric], "secondary": m, "info": info})
    except Exception as e:
        res.update({"status": "error", "score": None, "error": f"{type(e).__name__}: {e}",
                    "traceback": traceback.format_exc()[-3000:]})
    res["elapsed_s"] = round(time.time() - t0, 2)
    return res
