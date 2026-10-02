"""Segment-level probe on a real relational DB: RelBench rel-hm ``user-churn`` (binary, AUROC).

Question: do frozen-TFM hidden states carry an *entity-level* label through a training-free kNN probe, the way they
carried the planted segments on the synthetic shop DB? Item-level retrieval (user-item-purchase) is closed (negative);
this is the remaining, segment-level version of H1 on real data.

Protocol: customers labelled at the last training timestamp T0 (subsampled), features from the transactions in the
``hist_days`` before T0 (same customer_table as the purchase task), stratified K-fold over customers; each embedding is
scored by a kNN probe (mean neighbour label, cosine, k neighbours) and the references are the majority prior, HistGB on
the aggregates (supervised, model-free of TFMs) and TabPFN *supervised* on the aggregates (what the frozen model does
when it is given the labels in context instead of a pseudo-target).
"""
from __future__ import annotations

import time
from typing import Any

import numpy as np
import pandas as pd
from sklearn.metrics import roc_auc_score
from sklearn.model_selection import StratifiedKFold
from sklearn.neighbors import NearestNeighbors

from fer.relbench_hm import customer_table, embed_tabpfn, load_task, purchase_matrix, standardize, svd_features


def churn_context(task, db, hist_days: int = 365, n_customers: int | None = 20000, seed: int = 0):
    """Customers + churn labels at the last train timestamp, and the transactions in the window before it."""
    tbl = task.get_table("train", mask_input_cols=False).df
    t0 = pd.Timestamp(tbl["timestamp"].max())
    at = tbl[tbl["timestamp"] == t0].drop_duplicates("customer_id")
    if n_customers and len(at) > n_customers:
        at = at.sample(n_customers, random_state=seed)
    tx = db.table_dict["transactions"].df
    tx = tx[(tx["t_dat"] < t0) & (tx["t_dat"] >= t0 - pd.Timedelta(days=hist_days))]
    return at["customer_id"].to_numpy(), at["churn"].to_numpy().astype(int), t0, tx


def knn_probe(E: np.ndarray, y: np.ndarray, folds: list[tuple[np.ndarray, np.ndarray]], k: int = 50) -> float:
    En = E / (np.linalg.norm(E, axis=1, keepdims=True) + 1e-9)
    scores = np.zeros(len(y), dtype=np.float64)
    for tr, te in folds:
        nn = NearestNeighbors(n_neighbors=min(k, len(tr)), metric="cosine").fit(En[tr])
        _, idx = nn.kneighbors(En[te])
        scores[te] = y[tr][idx].mean(axis=1)
    return float(roc_auc_score(y, scores))


def supervised_probe(model: str, X: np.ndarray, y: np.ndarray, folds, device: str = "cuda") -> float:
    scores = np.zeros(len(y), dtype=np.float64)
    for tr, te in folds:
        if model == "hgb":
            from sklearn.ensemble import HistGradientBoostingClassifier

            clf = HistGradientBoostingClassifier(random_state=0).fit(X[tr], y[tr])
            scores[te] = clf.predict_proba(X[te])[:, 1]
        elif model == "tabpfn":
            import torch
            from tabpfn import TabPFNClassifier

            rng = np.random.default_rng(0)
            ctx = rng.choice(tr, min(len(tr), 3000), replace=False)  # in-context budget
            clf = TabPFNClassifier(device=device, n_estimators=4, inference_precision=torch.float32).fit(X[ctx], y[ctx])
            scores[te] = clf.predict_proba(X[te])[:, 1]
        else:
            raise KeyError(model)
    return float(roc_auc_score(y, scores))


def run_probe(embedders: list[str], hist_days: int = 365, n_customers: int | None = 20000, k: int = 50,
              n_folds: int = 5, seed: int = 0, device: str = "cuda", log: Any = None) -> dict[str, Any]:
    task, db = load_task("rel-hm", "user-churn")
    customers, y, t0, tx = churn_context(task, db, hist_days, n_customers, seed)
    art = db.table_dict["article"].df
    cust = db.table_dict["customer"].df
    row, agg = customer_table(tx, cust, art, customers, t0)
    feats = {"row": standardize(row.to_numpy()), "agg": standardize(agg.to_numpy())}
    folds = list(StratifiedKFold(n_folds, shuffle=True, random_state=seed).split(np.zeros(len(y)), y))
    res: dict[str, Any] = {"task": "rel-hm/user-churn", "t0": str(t0.date()), "n_customers": int(len(y)),
                           "pos_rate": float(y.mean()), "hist_days": hist_days, "k": k, "n_folds": n_folds, "rows": {}}

    def record(name, auroc, t1):
        res["rows"][name] = {"auroc": auroc, "seconds": round(time.time() - t1, 1)}
        if log is not None:
            log.event("churn_row", name=name, auroc=auroc)
        print(f"[churn] {name:36s} AUROC={auroc:.4f}", flush=True)

    t1 = time.time(); record("MajorityPrior", 0.5, t1)
    for ref in ("hgb", "tabpfn"):
        t1 = time.time()
        try:
            record(f"Supervised[{ref} on agg]", supervised_probe(ref, feats["agg"], y, folds, device), t1)
        except Exception as e:
            res["rows"][f"Supervised[{ref} on agg]"] = {"error": f"{type(e).__name__}: {e}"}
    P = None
    for name in embedders:
        t1 = time.time()
        try:
            if name in ("svd", "svd_agg", "tabpfn_svd_kmeans", "tabpfn_svd_random") and "svd" not in feats:
                P = purchase_matrix(tx, customers, art["article_id"].to_numpy())
                feats["svd"] = svd_features(P, seed=seed)
                feats["svd_agg"] = np.hstack([feats["svd"], feats["agg"]])
            if name in feats:
                E = feats[name]
            elif name.startswith("tabpfn_"):
                _, src, tgt = name.split("_", 2)
                E = embed_tabpfn(feats[src if src != "svd" else "svd_agg"], tgt, seed, device=device)
            else:
                raise KeyError(name)
            record(f"kNN[{name}]", knn_probe(E, y, folds, k), t1)
        except Exception as e:
            print(f"[churn] {name} FAILED: {type(e).__name__}: {e}", flush=True)
            res["rows"][f"kNN[{name}]"] = {"error": f"{type(e).__name__}: {e}"}
    return res
