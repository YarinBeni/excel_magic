"""Shared table helpers: ordinal encoding of non-numeric columns fitted on the training part (unknown -> -1), so the
same frames feed LightGBM / HGB and the tabular foundation models."""
from __future__ import annotations

import numpy as np
import pandas as pd


def encode_frames(train: pd.DataFrame, *others: pd.DataFrame, max_card: int = 10_000) -> list[pd.DataFrame]:
    out_tr = pd.DataFrame(index=train.index)
    outs = [pd.DataFrame(index=o.index) for o in others]
    for c in train.columns:
        s = train[c]
        if pd.api.types.is_bool_dtype(s) or pd.api.types.is_numeric_dtype(s):
            out_tr[c] = s.astype(float)
            for o, oo in zip(others, outs):
                oo[c] = o[c].astype(float)
            continue
        if pd.api.types.is_datetime64_any_dtype(s):
            out_tr[c] = s.astype("int64") / 86_400e9
            for o, oo in zip(others, outs):
                oo[c] = pd.to_datetime(o[c]).astype("int64") / 86_400e9
            continue
        vals = s.astype(str)
        cats = vals.value_counts().index[:max_card]
        m = {v: i for i, v in enumerate(cats)}
        out_tr[c] = vals.map(m).fillna(-1).astype(float)
        for o, oo in zip(others, outs):
            oo[c] = o[c].astype(str).map(m).fillna(-1).astype(float)
    return [out_tr, *outs]


def task_type(y: pd.Series) -> str:
    if pd.api.types.is_float_dtype(y) and y.nunique() > 20:
        return "regression"
    return "binary" if y.nunique() <= 2 else "multiclass"


def encode_target(y_tr: pd.Series, *ys: pd.Series) -> tuple[list[np.ndarray], list]:
    classes = sorted(pd.unique(y_tr.astype(str)))
    m = {c: i for i, c in enumerate(classes)}
    return [y_tr.astype(str).map(m).to_numpy()] + [y.astype(str).map(m).fillna(-1).astype(int).to_numpy() for y in ys], classes
