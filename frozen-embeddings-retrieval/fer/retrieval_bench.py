"""Retrieval benchmark for entity embeddings.

Task A  segment retrieval (synthetic DB only): for each customer, rank all other customers by cosine
        similarity; relevant = same latent segment (hidden ground truth). P@10, MAP@10, 10-NN accuracy.
Task B  future-purchase retrieval (kNN collaborative filtering): score products for each customer by
        the similarity-weighted pre-cutoff purchase counts of its K nearest customers; relevant =
        products the customer buys after the cutoff. MAP@10, recall@10, hit@10. Baselines: popularity.
"""
from __future__ import annotations

from typing import Any

import numpy as np
import pandas as pd

from .db import ShopDB


def _cosine(E: np.ndarray) -> np.ndarray:
    E = np.asarray(E, np.float32)
    E = E / np.clip(np.linalg.norm(E, axis=1, keepdims=True), 1e-9, None)
    return E @ E.T


def _ap_at_k(ranked_rel: np.ndarray, n_rel: int, k: int) -> float:
    if n_rel == 0:
        return np.nan
    hits = ranked_rel[:k]
    if hits.sum() == 0:
        return 0.0
    prec = np.cumsum(hits) / (np.arange(k) + 1)
    return float((prec * hits).sum() / min(n_rel, k))


def segment_retrieval(E: np.ndarray, segment: np.ndarray, k: int = 10) -> dict[str, float]:
    S = _cosine(E)
    np.fill_diagonal(S, -np.inf)
    n = len(S)
    topk = np.argpartition(-S, k, axis=1)[:, :k]
    # sort the top-k by similarity
    order = np.argsort(-np.take_along_axis(S, topk, 1), axis=1)
    topk = np.take_along_axis(topk, order, 1)
    rel = segment[topk] == segment[:, None]
    n_rel = pd.Series(segment).map(pd.Series(segment).value_counts()).to_numpy() - 1
    p_at_k = rel.mean()
    ap = np.array([_ap_at_k(rel[i], n_rel[i], k) for i in range(n)])
    knn_pred = np.array([np.bincount(segment[topk[i]], minlength=segment.max() + 1).argmax() for i in range(n)])
    chance = float(((pd.Series(segment).value_counts() - 1) / (n - 1) * pd.Series(segment).value_counts() / n).sum())
    return {"P@10": float(p_at_k), "MAP@10": float(np.nanmean(ap)), "kNN_acc": float((knn_pred == segment).mean()),
            "chance_P@10": chance}


def future_purchase_retrieval(E: np.ndarray, db: ShopDB, k_neighbors: int = 20, k: int = 10,
                              novel_only: bool = False) -> dict[str, float]:
    """novel_only=True scores only products the customer has NOT bought before the cutoff (candidates and truth):
    the harder "what new thing will they buy" task, immune to repeat-purchase saturation (Northwind)."""
    cust_ids = db.customer_ids()
    cidx = {c: i for i, c in enumerate(cust_ids)}
    prod_ids = db.products.product_id.to_numpy()
    pidx = {p: i for i, p in enumerate(prod_ids)}
    past, fut = db.past_items(), db.future_items()
    P = np.zeros((len(cust_ids), len(prod_ids)), np.float32)  # pre-cutoff purchase counts
    np.add.at(P, (past.customer_id.map(cidx).to_numpy(), past.product_id.map(pidx).to_numpy()), past.quantity.to_numpy(float))
    truth: dict[int, set[int]] = {}
    for c, p in zip(fut.customer_id.map(cidx), fut.product_id.map(pidx)):
        if novel_only and P[int(c), int(p)] > 0:
            continue
        truth.setdefault(int(c), set()).add(int(p))
    queries = sorted(truth)
    S = _cosine(E)
    np.fill_diagonal(S, -np.inf)
    pop = P.sum(0)
    res: dict[str, list[float]] = {"MAP@10": [], "recall@10": [], "hit@10": [], "pop_MAP@10": [], "pop_recall@10": []}
    for q in queries:
        nb = np.argpartition(-S[q], k_neighbors)[:k_neighbors]
        w = np.clip(S[q, nb], 0, None)
        scores = (w[:, None] * P[nb]).sum(0) + 1e-6 * pop  # tie-break by popularity
        pop_q = pop.copy()
        if novel_only:  # already-bought products are not candidates
            scores = np.where(P[q] > 0, -np.inf, scores)
            pop_q = np.where(P[q] > 0, -np.inf, pop_q)
        ranked = np.argsort(-scores)[:k]
        rel = np.array([r in truth[q] for r in ranked])
        res["MAP@10"].append(_ap_at_k(rel, len(truth[q]), k))
        res["recall@10"].append(rel.sum() / len(truth[q]))
        res["hit@10"].append(float(rel.any()))
        ranked_pop = np.argsort(-pop_q)[:k]
        relp = np.array([r in truth[q] for r in ranked_pop])
        res["pop_MAP@10"].append(_ap_at_k(relp, len(truth[q]), k))
        res["pop_recall@10"].append(relp.sum() / len(truth[q]))
    out = {m: float(np.nanmean(v)) for m, v in res.items()}
    out["n_queries"] = len(queries)
    return out


def run_benchmark(E: np.ndarray, db: ShopDB, k_neighbors: int = 20) -> dict[str, Any]:
    out: dict[str, Any] = {"dim": int(E.shape[1])}
    if db.hidden is not None and "segment" in db.hidden.columns:
        seg = db.hidden.set_index("customer_id").segment.reindex(db.customer_ids()).to_numpy()
        out["segment"] = segment_retrieval(E, seg)
    out["future_purchase"] = future_purchase_retrieval(E, db, k_neighbors=k_neighbors)
    out["future_purchase_novel"] = future_purchase_retrieval(E, db, k_neighbors=k_neighbors, novel_only=True)
    return out
