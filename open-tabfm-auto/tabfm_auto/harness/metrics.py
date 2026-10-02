"""Primary metric per task type follows TabArena / TabFM-Auto: lower is better.

binary -> 1 - AUROC, multiclass -> log loss, regression -> RMSE.
"""
from __future__ import annotations

import numpy as np
from sklearn.metrics import (
    accuracy_score,
    log_loss,
    mean_absolute_error,
    r2_score,
    roc_auc_score,
    root_mean_squared_error,
)


def primary_metric_name(task_type: str) -> str:
    return {"binary": "1-auroc", "multiclass": "logloss", "regression": "rmse"}[task_type]


def score(task_type: str, y_true, pred, classes=None) -> dict[str, float]:
    y_true = np.asarray(y_true)
    pred = np.asarray(pred, dtype=float)
    if task_type == "regression":
        return {"rmse": float(root_mean_squared_error(y_true, pred)),
                "mae": float(mean_absolute_error(y_true, pred)),
                "r2": float(r2_score(y_true, pred))}
    classes = np.asarray(classes) if classes is not None else np.unique(y_true)
    pred = np.clip(pred, 1e-9, 1)
    pred = pred / pred.sum(1, keepdims=True)
    yhat = classes[pred.argmax(1)]
    out = {"logloss": float(log_loss(y_true, pred, labels=classes)),
           "acc": float(accuracy_score(y_true, yhat))}
    if task_type == "binary":
        pos = 1 if len(classes) == 2 else 0
        out["auroc"] = float(roc_auc_score(y_true == classes[pos], pred[:, pos]))
        out["1-auroc"] = 1.0 - out["auroc"]
    return out
