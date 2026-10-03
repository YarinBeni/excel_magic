"""The DEEP tool: frozen tabular / relational foundation models as a statistician the LLM can call (capabilities T1-T7).

Every method takes pandas frames that SQL produced (the LLM never sees the rows) and returns a small JSON-able dict.
Backbones are spec strings of open-tabfm-auto's registry, e.g. "kumo-tabular-l:n_estimators=4,device=cuda" (accurate),
"kumo-tabular-s:n_estimators=2,device=cuda" (fast, for loops), "lightgbm" (the check on temporal / large tasks).

  T1 predict      held-out quality (time split when a time column is given) + per-entity scores for query rows
  T2 drivers      permutation importance on held-out rows, mean and spread over repeats, with the GBDT's ranking
  T3 what_if      partial dependence of the prediction on one column (mean and 10-90% band)
  T4 anomalies    out-of-fold surprise per row: -log p(true class) or |residual| / MAD
  T5 drift        classifier two-sample test between two periods: AUROC (0.5 = no change) and which columns moved
  T6 similar      nearest entities by cosine over an embedding matrix (e.g. Kumo Relational's graph layer)
  T7 hypothesis   held-out gain of adding columns, paired over repeated splits: delta, standard error, z
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any

import numpy as np
import pandas as pd

from .tabular import encode_frames, encode_target, task_type


def _registry():
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "open-tabfm-auto"))
    from tabfm_auto.models.registry import get_model

    return get_model


def _metric(task: str, y, p) -> tuple[str, float]:
    from sklearn.metrics import accuracy_score, r2_score, roc_auc_score

    if task == "binary":
        return "auroc", float(roc_auc_score(y, p[:, 1])) if len(set(y)) == 2 else float("nan")
    if task == "multiclass":
        return "accuracy", float(accuracy_score(y, p.argmax(1)))
    return "r2", float(r2_score(y, p))


@dataclass
class DeepTool:
    model: str = "kumo-tabular-l:n_estimators=4,device=cuda"
    fast: str = "kumo-tabular-s:n_estimators=2,device=cuda"
    check: str = "lightgbm"
    max_context: int = 10_000
    holdout: float = 0.2
    seed: int = 0

    # ----------------------------------------------------------------------------------------------- internals
    def _split(self, df: pd.DataFrame, time_col: str | None, rng) -> tuple[np.ndarray, np.ndarray]:
        n = len(df)
        if time_col:
            order = np.argsort(pd.to_datetime(df[time_col]).to_numpy(), kind="stable")
            cut = int(n * (1 - self.holdout))
            return order[:cut], order[cut:]
        perm = rng.permutation(n)
        cut = int(n * (1 - self.holdout))
        return perm[:cut], perm[cut:]

    def _fit_predict(self, spec: str, task: str, Xtr, ytr, Xte, rng) -> np.ndarray:
        if len(Xtr) > self.max_context and not spec.startswith(("lightgbm", "hgb", "rf")):
            idx = rng.choice(len(Xtr), self.max_context, replace=False)
            Xtr, ytr = Xtr.iloc[idx], ytr[idx]
        est = _registry()(spec, task)
        est.fit(Xtr, ytr)
        if task == "regression":
            return np.asarray(est.predict(Xte), float)
        p = np.asarray(est.predict_proba(Xte), float)
        if task == "binary" and p.ndim == 2 and p.shape[1] == 1:
            p = np.hstack([1 - p, p])
        return p

    def _prep(self, df: pd.DataFrame, target: str, drop: list[str]) -> tuple[pd.DataFrame, np.ndarray, str, list]:
        y_raw = df[target]
        task = task_type(y_raw)
        X = df.drop(columns=[target] + [c for c in drop if c in df.columns])
        if task == "regression":
            return X, y_raw.astype(float).to_numpy(), task, []
        (y,), classes = encode_target(y_raw)
        return X, y, task, classes

    # ----------------------------------------------------------------------------------------------- T1
    def predict(self, df: pd.DataFrame, target: str, time_col: str | None = None, id_col: str | None = None,
                query: pd.DataFrame | None = None, top: int = 20) -> dict[str, Any]:
        rng = np.random.default_rng(self.seed)
        drop = [c for c in (time_col, id_col) if c]
        X, y, task, classes = self._prep(df, target, drop)
        tr, te = self._split(df, time_col, rng)
        Xtr, Xte = encode_frames(X.iloc[tr], X.iloc[te])
        t0 = time.time()
        p = self._fit_predict(self.model, task, Xtr, y[tr], Xte, rng)
        name, val = _metric(task, y[te], p)
        out = {"task": task, "metric": name, "value": round(val, 4), "n_train": len(tr), "n_holdout": len(te),
               "split": "time" if time_col else "random", "model": self.model.split(":")[0],
               "seconds": round(time.time() - t0, 1)}
        if self.check and (time_col or len(df) > 100_000):
            _, cval = _metric(task, y[te], self._fit_predict(self.check, task, Xtr, y[tr], Xte, rng))
            out["check_model"], out["check_value"] = self.check, round(cval, 4)
        if query is not None:
            Xall, Xq = encode_frames(X, query.drop(columns=[c for c in [target] + drop if c in query.columns]))
            pq = self._fit_predict(self.model, task, Xall, y, Xq, rng)
            score = pq[:, 1] if task == "binary" else pq if task == "regression" else pq.max(1)
            order = np.argsort(-score)[:top]
            ids = query[id_col].to_numpy()[order] if id_col else order
            out["query_top"] = [{"id": (i.item() if hasattr(i, "item") else i), "score": round(float(score[j]), 4)}
                                for i, j in zip(ids, order)]
            if classes:
                out["classes"] = classes
        return out

    # ----------------------------------------------------------------------------------------------- T2
    def drivers(self, df: pd.DataFrame, target: str, time_col: str | None = None, id_col: str | None = None,
                top: int = 10, repeats: int = 3, fast: bool = True) -> dict[str, Any]:
        rng = np.random.default_rng(self.seed)
        drop = [c for c in (time_col, id_col) if c]
        X, y, task, _ = self._prep(df, target, drop)
        tr, te = self._split(df, time_col, rng)
        Xtr, Xte = encode_frames(X.iloc[tr], X.iloc[te])
        spec = self.fast if fast else self.model
        rng2 = np.random.default_rng(self.seed)
        sub = rng2.choice(len(Xtr), min(len(Xtr), self.max_context), replace=False)
        est = _registry()(spec, task)
        est.fit(Xtr.iloc[sub], y[tr][sub])

        def score(Xe):
            p = est.predict(Xe) if task == "regression" else est.predict_proba(Xe)
            return _metric(task, y[te], np.asarray(p, float))[1]

        base = score(Xte)
        imp = {}
        for c in Xte.columns:
            drops = []
            for _ in range(repeats):
                Xp = Xte.copy()
                Xp[c] = rng.permutation(Xp[c].to_numpy())
                drops.append(base - score(Xp))
            imp[c] = (float(np.mean(drops)), float(np.std(drops)))
        ranked = sorted(imp.items(), key=lambda kv: -kv[1][0])
        out = {"task": task, "base_value": round(base, 4), "model": spec.split(":")[0],
               "drivers": [{"column": c, "importance": round(m, 4), "spread": round(s, 4)} for c, (m, s) in ranked[:top]]}
        if self.check:
            ce = _registry()(self.check, task)
            ce.fit(Xtr, y[tr])
            gi = getattr(ce, "feature_importances_", None)
            if gi is None and hasattr(ce, "model_"):
                gi = getattr(ce.model_, "feature_importances_", None)
            if gi is not None:
                order = np.argsort(-np.asarray(gi))[:top]
                out["check_ranking"] = [str(Xtr.columns[i]) for i in order]
        return out

    # ----------------------------------------------------------------------------------------------- T3
    def what_if(self, df: pd.DataFrame, target: str, column: str, values: list | None = None,
                time_col: str | None = None, id_col: str | None = None, n_rows: int = 500) -> dict[str, Any]:
        rng = np.random.default_rng(self.seed)
        drop = [c for c in (time_col, id_col) if c]
        X, y, task, classes = self._prep(df, target, drop)
        if values is None:
            s = X[column]
            values = (list(np.unique(np.quantile(s.dropna(), [0.05, 0.25, 0.5, 0.75, 0.95])))
                      if pd.api.types.is_numeric_dtype(s) else list(s.astype(str).value_counts().index[:8]))
        rows = rng.choice(len(X), min(n_rows, len(X)), replace=False)
        grid = []
        for v in values:
            Xv = X.iloc[rows].copy()
            Xv[column] = v
            grid.append(Xv)
        Xall, *Xg = encode_frames(X, *grid)
        # one fit, many predictions
        sub = rng.choice(len(Xall), min(len(Xall), self.max_context), replace=False)
        est = _registry()(self.model, task)
        est.fit(Xall.iloc[sub], y[sub])
        out = []
        for v, Xv in zip(values, Xg):
            p = np.asarray(est.predict(Xv) if task == "regression" else est.predict_proba(Xv), float)
            s = p if task == "regression" else p[:, 1] if task == "binary" else p.max(1)
            out.append({"value": v if not hasattr(v, "item") else v.item(), "mean": round(float(s.mean()), 4),
                        "p10": round(float(np.quantile(s, 0.1)), 4), "p90": round(float(np.quantile(s, 0.9)), 4)})
        return {"task": task, "column": column, "n_rows": len(rows), "model": self.model.split(":")[0], "curve": out,
                **({"positive_class": classes[1]} if task == "binary" and classes else {})}

    # ----------------------------------------------------------------------------------------------- T4
    def anomalies(self, df: pd.DataFrame, target: str, id_col: str | None = None, time_col: str | None = None,
                  top: int = 20, folds: int = 2) -> dict[str, Any]:
        rng = np.random.default_rng(self.seed)
        drop = [c for c in (time_col, id_col) if c]
        X, y, task, _ = self._prep(df, target, drop)
        f = rng.permutation(len(X)) % folds
        surprise = np.zeros(len(X))
        for k in range(folds):
            tr, te = np.where(f != k)[0], np.where(f == k)[0]
            Xtr, Xte = encode_frames(X.iloc[tr], X.iloc[te])
            p = self._fit_predict(self.fast, task, Xtr, y[tr], Xte, rng)
            if task == "regression":
                r = np.abs(y[te] - p)
                surprise[te] = r / (np.median(np.abs(y[tr] - np.median(y[tr]))) + 1e-9)
            else:
                surprise[te] = -np.log(np.clip(p[np.arange(len(te)), y[te]], 1e-9, 1))
        order = np.argsort(-surprise)[:top]
        ids = df[id_col].to_numpy()[order] if id_col else order
        return {"task": task, "model": self.fast.split(":")[0], "n_rows": len(X),
                "surprise_p50": round(float(np.median(surprise)), 4), "surprise_p99": round(float(np.quantile(surprise, 0.99)), 4),
                "top": [{"id": (i.item() if hasattr(i, "item") else i), "surprise": round(float(surprise[j]), 4)}
                        for i, j in zip(ids, order)]}

    # ----------------------------------------------------------------------------------------------- T5
    def drift(self, old: pd.DataFrame, new: pd.DataFrame, exclude: list[str] | None = None, top: int = 10,
              folds: int = 2) -> dict[str, Any]:
        cols = [c for c in old.columns if c in new.columns and c not in (exclude or [])]
        df = pd.concat([old[cols].assign(_period=0), new[cols].assign(_period=1)], ignore_index=True)
        rng = np.random.default_rng(self.seed)
        if len(df) > 2 * self.max_context:
            df = df.iloc[rng.choice(len(df), 2 * self.max_context, replace=False)].reset_index(drop=True)
        X, y = df[cols], df["_period"].to_numpy()
        f = rng.permutation(len(X)) % folds
        p = np.zeros(len(X))
        from sklearn.metrics import roc_auc_score

        imp = {c: [] for c in cols}
        for k in range(folds):
            tr, te = np.where(f != k)[0], np.where(f == k)[0]
            Xtr, Xte = encode_frames(X.iloc[tr], X.iloc[te])
            est = _registry()(self.fast, "binary")
            est.fit(Xtr, y[tr])
            p[te] = np.asarray(est.predict_proba(Xte))[:, 1]
            base = roc_auc_score(y[te], p[te])
            for c in cols:
                Xp = Xte.copy()
                Xp[c] = rng.permutation(Xp[c].to_numpy())
                imp[c].append(base - roc_auc_score(y[te], np.asarray(est.predict_proba(Xp))[:, 1]))
        auc = float(roc_auc_score(y, p))
        ranked = sorted(((c, float(np.mean(v))) for c, v in imp.items()), key=lambda kv: -kv[1])
        return {"auroc": round(auc, 4), "changed": auc > 0.55, "n_old": int(len(old)), "n_new": int(len(new)),
                "model": self.fast.split(":")[0],
                "columns": [{"column": c, "importance": round(v, 4)} for c, v in ranked[:top]]}

    # ----------------------------------------------------------------------------------------------- T6
    @staticmethod
    def similar(emb: np.ndarray, ids: list, query_ids: list, k: int = 10) -> dict[str, Any]:
        E = emb / (np.linalg.norm(emb, axis=1, keepdims=True) + 1e-9)
        pos = {v: i for i, v in enumerate(ids)}
        q = np.array([pos[i] for i in query_ids if i in pos])
        centroid = E[q].mean(0)
        sims = E @ centroid
        sims[q] = -np.inf
        order = np.argsort(-sims)[:k]
        return {"query_size": int(len(q)), "neighbours": [{"id": ids[j], "cosine": round(float(sims[j]), 4)} for j in order]}

    # ----------------------------------------------------------------------------------------------- T7
    def hypothesis(self, df: pd.DataFrame, target: str, columns: list[str], time_col: str | None = None,
                   id_col: str | None = None, repeats: int = 5) -> dict[str, Any]:
        drop = [c for c in (time_col, id_col) if c]
        X, y, task, _ = self._prep(df, target, drop)
        deltas = []
        for r in range(repeats):
            rng = np.random.default_rng(self.seed + r)
            tr, te = self._split(df, None, rng)  # repeated random splits: a paired comparison on the same rows
            Xtr, Xte = encode_frames(X.iloc[tr], X.iloc[te])
            with_ = _metric(task, y[te], self._fit_predict(self.fast, task, Xtr, y[tr], Xte, rng))[1]
            keep = [c for c in Xtr.columns if c not in columns]
            without = _metric(task, y[te], self._fit_predict(self.fast, task, Xtr[keep], y[tr], Xte[keep], rng))[1]
            deltas.append(with_ - without)
        d = np.array(deltas)
        se = float(d.std(ddof=1) / np.sqrt(len(d))) if len(d) > 1 else float("nan")
        return {"task": task, "columns": columns, "delta_mean": round(float(d.mean()), 4), "delta_se": round(se, 4),
                "z": round(float(d.mean() / se), 2) if se and np.isfinite(se) and se > 0 else None,
                "supported": bool(se and d.mean() > 2 * se), "repeats": repeats, "model": self.fast.split(":")[0]}
