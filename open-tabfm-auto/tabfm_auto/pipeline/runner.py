"""Execute a candidate pipeline around a frozen model for one train/test split."""
from __future__ import annotations

import time
import types
from typing import Any

import numpy as np
import pandas as pd

from ..models import get_model
from .loader import MAX_COLUMNS, PipelineError

# model constructor keys the pipeline may override (everything else is dropped with a warning)
ALLOWED_MODEL_KWARGS = {"n_estimators", "softmax_temperature", "balance_probabilities",
                        "average_before_softmax", "random_state", "inference_precision",
                        "fit_mode", "memory_saving_mode", "n_preprocessing_jobs", "differentiable_input"}


def _as_frame(X: Any, like: pd.DataFrame | None = None) -> pd.DataFrame:
    if isinstance(X, pd.DataFrame):
        return X
    X = np.asarray(X)
    cols = list(like.columns) if like is not None and like.shape[1] == X.shape[1] else [f"f{i}" for i in range(X.shape[1])]
    return pd.DataFrame(X, columns=cols)


def _sanitize(X: pd.DataFrame) -> pd.DataFrame:
    """Make a frame consumable by TabPFN / sklearn: numeric or category dtypes, finite floats."""
    X = X.copy()
    for c in X.columns:
        s = X[c]
        if pd.api.types.is_bool_dtype(s):
            X[c] = s.astype(float)
        elif pd.api.types.is_numeric_dtype(s):
            X[c] = pd.to_numeric(s, errors="coerce").astype(float).replace([np.inf, -np.inf], np.nan)
        elif isinstance(s.dtype, pd.CategoricalDtype):
            X[c] = s.cat.codes.astype(float).replace(-1, np.nan)
        else:
            X[c] = pd.factorize(s.astype(str), use_na_sentinel=True)[0].astype(float)
            X.loc[s.isna(), c] = np.nan
    X.columns = [str(c) for c in X.columns]
    return X


def _seed_everything(seed: int) -> None:
    import random

    random.seed(seed)
    np.random.seed(seed)
    try:
        import torch

        torch.manual_seed(seed)
        if torch.cuda.is_available():
            torch.cuda.manual_seed_all(seed)
    except Exception:  # torch absent (classical backbones only)
        pass


def run_pipeline(mod: types.ModuleType, X_train: pd.DataFrame, y_train: pd.Series, X_test: pd.DataFrame,
                 task_type: str, model_spec: str = "tabpfn", max_rows: int = 10000, seed: int = 0,
                 log: Any = None) -> tuple[np.ndarray, dict[str, Any]]:
    """Returns (predictions, info). predictions: (n_test, n_classes) probabilities or (n_test,) values."""
    info: dict[str, Any] = {}
    t0 = time.time()
    X_train, X_test = X_train.reset_index(drop=True), X_test.reset_index(drop=True)
    y_train = pd.Series(np.asarray(y_train), name=y_train.name or "target")
    y_orig = y_train.copy()
    classes = np.unique(y_orig) if task_type != "regression" else None
    _seed_everything(seed)  # pipelines and backbones (Kumo-S varies +-30% per split unseeded, J16) are reproducible per seed

    # 1. preprocess
    out = mod.preprocess(X_train.copy(), y_train.copy(), X_test.copy())
    if not (isinstance(out, (tuple, list)) and len(out) == 3):
        raise PipelineError("preprocess must return (X_train, y_train, X_test)")
    X_tr, y_tr, X_te = _as_frame(out[0], X_train), pd.Series(np.asarray(out[1])), _as_frame(out[2], X_test)
    if len(X_tr) != len(y_tr):
        raise PipelineError("preprocess returned X_train / y_train with different lengths")
    if len(X_te) != len(X_test):
        raise PipelineError("preprocess must not change the number of test rows")

    # 2. engineer
    out = mod.engineer(X_tr.copy(), y_tr.copy(), X_te.copy())
    if not (isinstance(out, (tuple, list)) and len(out) == 2):
        raise PipelineError("engineer must return (X_train, X_test)")
    X_tr, X_te = _as_frame(out[0], X_tr), _as_frame(out[1], X_te)
    if list(X_tr.columns) != list(X_te.columns):
        raise PipelineError("engineer returned different columns for train and test")
    if X_tr.shape[1] > MAX_COLUMNS:
        raise PipelineError(f"engineer returned {X_tr.shape[1]} columns (> {MAX_COLUMNS})")
    if X_tr.shape[1] == 0:
        raise PipelineError("engineer returned zero columns")
    X_tr, X_te = _sanitize(X_tr), _sanitize(X_te)
    info["n_features"] = int(X_tr.shape[1])
    info["t_features"] = round(time.time() - t0, 3)

    # 3. sample -> context views
    views = mod.sample(X_tr.copy(), y_tr.copy(), X_te.copy(), max_rows)
    if isinstance(views, np.ndarray) and views.ndim == 1:
        views = [views]
    if not isinstance(views, (list, tuple)) or len(views) == 0:
        raise PipelineError("sample must return a non-empty list of index arrays")
    views = [np.asarray(v, dtype=int) for v in views]
    for v in views:
        if len(v) == 0 or v.min() < 0 or v.max() >= len(X_tr):
            raise PipelineError("sample returned invalid indices")
        if len(v) > max_rows:
            raise PipelineError(f"a context view has {len(v)} rows (> max_rows={max_rows})")
    info["n_views"] = len(views)
    info["view_sizes"] = [int(len(v)) for v in views]

    # 4. frozen model, once per view
    mk = {k: v for k, v in getattr(mod, "MODEL_KWARGS", {}).items() if k in ALLOWED_MODEL_KWARGS}
    dropped = sorted(set(getattr(mod, "MODEL_KWARGS", {})) - set(mk))
    if dropped and log is not None:
        log.warning("MODEL_KWARGS keys ignored: %s", dropped)
    spec = model_spec + ("," if ":" in model_spec else ":") + ",".join(f"{k}={v}" for k, v in mk.items()) if mk else model_spec
    t1 = time.time()
    preds = []
    for i, idx in enumerate(views):
        model = get_model(spec, task_type)
        if task_type != "regression":
            yv = y_tr.iloc[idx]
            if yv.nunique() < 2:
                raise PipelineError(f"context view {i} contains a single class")
        model.fit(X_tr.iloc[idx], y_tr.iloc[idx])
        if task_type == "regression":
            p = np.asarray(model.predict(X_te), dtype=float).reshape(-1)
        else:
            proba = np.asarray(model.predict_proba(X_te), dtype=float)
            # align to the global class set (a view may miss classes)
            mcls = np.asarray(getattr(model, "classes_", np.unique(y_tr.iloc[idx])))
            p = np.zeros((len(X_te), len(classes)))
            for j, c in enumerate(mcls):
                k = np.where(classes == c)[0]
                if len(k):
                    p[:, k[0]] = proba[:, j]
            p = p / np.clip(p.sum(1, keepdims=True), 1e-12, None)
        preds.append(p)
    pred = np.mean(preds, axis=0)
    info["t_model"] = round(time.time() - t1, 3)

    # 5. postprocess
    if task_type == "regression":
        pred_obj: Any = pd.Series(pred, name="prediction")
    else:
        pred_obj = pd.DataFrame(pred, columns=list(classes))
    out = mod.postprocess(pred_obj, y_orig.copy(), X_te.copy())
    if task_type == "regression":
        final = np.asarray(out, dtype=float).reshape(-1)
        if final.shape[0] != len(X_test):
            raise PipelineError("postprocess changed the number of predictions")
    else:
        if isinstance(out, pd.DataFrame):
            out = out.reindex(columns=list(classes))
        final = np.asarray(out, dtype=float)
        if final.shape != (len(X_test), len(classes)):
            raise PipelineError(f"postprocess must return shape {(len(X_test), len(classes))}, got {final.shape}")
        final = np.clip(final, 1e-9, None)
        final = final / final.sum(1, keepdims=True)
    if not np.all(np.isfinite(final)):
        raise PipelineError("non-finite predictions")
    info["t_total"] = round(time.time() - t0, 3)
    info["model_spec"] = spec
    return final, info
