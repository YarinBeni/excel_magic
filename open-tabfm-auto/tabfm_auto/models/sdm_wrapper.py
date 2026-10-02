"""sklearn-style wrapper around NVIDIA structured-data-models (``sdm``) in-context learners
(KumoTabular S/M/L, TabICLv2). Features must be numeric (the pipeline runner guarantees that)."""
from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd


class SdmEstimator:
    def __init__(self, family: str, task_type: str, size: str | None = None, n_estimators: int = 4,
                 device: str = "cpu", **_: Any):
        import sdm
        import torch

        self.task_type = task_type
        self.n_estimators = n_estimators
        self.device = torch.device(device)
        task = "regression" if task_type == "regression" else "classification"
        if family == "kumo-tabular":
            self.model = sdm.models.KumoTabular(task=task, size=size or "small", device=self.device)
        elif family == "tabiclv2":
            self.model = sdm.models.TabICLv2(task=task, device=self.device)
        else:
            raise KeyError(family)
        self.model.eval()
        self._sdm = sdm

    def _table(self, X: pd.DataFrame, y: pd.Series | None = None):
        df = pd.DataFrame(np.asarray(X, dtype=np.float32), columns=[str(c) for c in X.columns])
        stypes: dict[str, str] = {c: "numerical" for c in df.columns}
        if y is not None:
            if self.task_type == "regression":
                df["__target__"] = np.asarray(y, dtype=np.float32)
                stypes["__target__"] = "numerical"
            else:
                df["__target__"] = pd.Categorical(np.asarray(y).astype(str))
                stypes["__target__"] = "categorical"
        return self._sdm.TableTensor.from_pandas(df=df, stypes=stypes, device=self.device)

    def fit(self, X: pd.DataFrame, y: pd.Series):
        import torch

        y = pd.Series(np.asarray(y))
        if self.task_type != "regression":
            self.classes_ = np.unique(y)
        t = self._table(X, y)
        with torch.no_grad():
            self.model.fit(x=t.drop_columns("__target__"), y=t[:, "__target__"], num_estimators=self.n_estimators)
        return self

    def _predict_frame(self, X: pd.DataFrame) -> pd.DataFrame:
        import torch

        with torch.no_grad():
            out = self.model.predict(self._table(X))
        return out.to_pandas()

    def predict_proba(self, X: pd.DataFrame) -> np.ndarray:
        df = self._predict_frame(X)
        cols = {str(c): c for c in df.columns}
        P = np.zeros((len(df), len(self.classes_)), dtype=float)
        for j, cls in enumerate(self.classes_):
            key = cols.get(str(cls))
            if key is not None:
                P[:, j] = df[key].to_numpy(dtype=float)
        return P / np.clip(P.sum(1, keepdims=True), 1e-12, None)

    def predict(self, X: pd.DataFrame) -> np.ndarray:
        if self.task_type != "regression":
            return self.classes_[self.predict_proba(X).argmax(1)]
        df = self._predict_frame(X)
        if "q500" in df.columns:
            return df["q500"].to_numpy(dtype=float)
        return df.select_dtypes("number").mean(axis=1).to_numpy(dtype=float)

    def clear(self) -> None:
        self.model.clear()
