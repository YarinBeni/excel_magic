"""RelBench rel-hm / user-item-purchase (MAP@12, 7-day window) with frozen-embedding kNN retrieval.

Published rows to reproduce (test MAP x100): GlobalPopularity 0.30, PastVisit 0.89, LightGBM 0.38, RDL-GraphSAGE 0.80,
ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93 (see docs/benchmark-candidates-2026-10-02.md; verify vs PDFs).

Our rows (training-free): customers embedded from their pre-cutoff history, articles ranked by user-kNN collaborative
filtering in embedding space (neighbours' purchases, cosine-weighted), optionally hybridised with PastVisit.
All scoring goes through ``task.evaluate`` so the numbers are the official ones.
"""
from __future__ import annotations

import time
from typing import Any

import numpy as np
import pandas as pd
import scipy.sparse as sp


# ------------------------------------------------------------------------------------------- data
def load_task(dataset: str = "rel-hm", task_name: str = "user-item-purchase"):
    import relbench

    ds = relbench.load_dataset(dataset)
    task = ds.load_task(task_name)
    db = task.get_db(upto_test_timestamp=False)
    return task, db


def split_context(task, db, split: str, hist_days: int = 365):
    """Query customers of ``split`` and the transactions strictly before the split timestamp (last ``hist_days``)."""
    tbl = task.get_table(split, mask_input_cols=False).df
    cutoff = pd.Timestamp(tbl["timestamp"].min())
    tx = db.table_dict["transactions"].df
    tx = tx[(tx["t_dat"] < cutoff) & (tx["t_dat"] >= cutoff - pd.Timedelta(days=hist_days))]
    queries = tbl["customer_id"].drop_duplicates().to_numpy()
    return tbl, cutoff, tx, queries


# ------------------------------------------------------------------------------------------- features
def customer_table(tx: pd.DataFrame, cust: pd.DataFrame, art: pd.DataFrame, customers: np.ndarray, cutoff: pd.Timestamp,
                   mix_col: str = "product_group_name") -> tuple[pd.DataFrame, pd.DataFrame]:
    """(row-only frame, aggregate frame) indexed by customer_id for ``customers``."""
    c = cust.set_index("customer_id").reindex(customers)
    row = pd.DataFrame(index=c.index)
    row["age"] = pd.to_numeric(c["age"], errors="coerce").fillna(-1)
    for col in ("club_member_status", "fashion_news_frequency", "Active", "FN"):
        if col in c.columns:
            row[col] = pd.factorize(c[col].astype(str))[0]
    t = tx[tx["customer_id"].isin(customers)].merge(art[["article_id", mix_col]], on="article_id", how="left")
    g = t.groupby("customer_id")
    agg = pd.DataFrame(index=row.index)
    agg["n_tx"] = g.size().reindex(agg.index).fillna(0)
    agg["n_articles"] = g["article_id"].nunique().reindex(agg.index).fillna(0)
    agg["spend"] = g["price"].sum().reindex(agg.index).fillna(0)
    agg["mean_price"] = g["price"].mean().reindex(agg.index).fillna(0)
    agg["days_since_last"] = (cutoff - g["t_dat"].max()).dt.days.reindex(agg.index).fillna(9999)
    agg["days_since_first"] = (cutoff - g["t_dat"].min()).dt.days.reindex(agg.index).fillna(9999)
    agg["n_tx_30d"] = t[t["t_dat"] >= cutoff - pd.Timedelta(days=30)].groupby("customer_id").size().reindex(agg.index).fillna(0)
    agg["n_tx_90d"] = t[t["t_dat"] >= cutoff - pd.Timedelta(days=90)].groupby("customer_id").size().reindex(agg.index).fillna(0)
    if "sales_channel_id" in t.columns:
        agg["online_share"] = (t["sales_channel_id"] == 2).groupby(t["customer_id"]).mean().reindex(agg.index).fillna(0)
    mix = pd.crosstab(t["customer_id"], t[mix_col].astype(str), normalize="index").reindex(agg.index).fillna(0)
    mix.columns = [f"mix_{x}" for x in mix.columns]
    return row, pd.concat([row, agg, mix], axis=1)


def standardize(M: np.ndarray) -> np.ndarray:
    M = np.asarray(M, np.float32)
    mu, sd = np.nanmean(M, 0), np.nanstd(M, 0)
    sd[sd == 0] = 1
    return np.nan_to_num((M - mu) / sd)


# ------------------------------------------------------------------------------------------- embedders
def embed_tabpfn(X: np.ndarray, target: str = "kmeans", seed: int = 0, n_context: int = 1000, chunk: int = 4096,
                 device: str = "cuda", n_clusters: int = 8) -> np.ndarray:
    """Frozen TabPFN v2 per-row hidden state over X; in-context target = k-means pseudo-labels (default) or random."""
    from sklearn.cluster import KMeans
    from tabfm_auto.models import get_model

    rng = np.random.default_rng(seed)
    n = len(X)
    Xdf = pd.DataFrame(X, columns=[f"f{i}" for i in range(X.shape[1])])
    if target == "kmeans":
        y = KMeans(n_clusters, n_init=4, random_state=seed).fit_predict(X)
        task_type = "multiclass"
    else:
        y = rng.integers(0, 2, n)
        task_type = "binary"
    ctx = rng.choice(n, min(n, n_context), replace=False)
    clf = get_model(f"tabpfn:n_estimators=1,device={device},inference_precision=float32", task_type)
    clf.fit(Xdf.iloc[ctx], y[ctx])
    out = []
    for s in range(0, n, chunk):
        E = clf.get_embeddings(Xdf.iloc[s:s + chunk], data_source="test")
        if hasattr(E, "detach"):
            E = E.detach().float().cpu().numpy()
        E = np.asarray(E, dtype=np.float32)
        out.append(E.mean(0) if E.ndim == 3 else E)
    return np.concatenate(out, 0)


# ------------------------------------------------------------------------------------------- retrieval
def purchase_matrix(tx: pd.DataFrame, customers: np.ndarray, articles: np.ndarray) -> sp.csr_matrix:
    cidx = pd.Series(np.arange(len(customers)), index=customers)
    aidx = pd.Series(np.arange(len(articles)), index=articles)
    t = tx[tx["customer_id"].isin(cidx.index) & tx["article_id"].isin(aidx.index)]
    return sp.csr_matrix((np.ones(len(t), np.float32), (cidx[t["customer_id"]].to_numpy(), aidx[t["article_id"]].to_numpy())),
                         shape=(len(customers), len(articles)))


def svd_features(P: sp.csr_matrix, n_components: int = 64, seed: int = 0) -> np.ndarray:
    """Dense customer features from the purchase matrix itself: truncated SVD of log1p counts, standardized.
    This is the interaction (graph) signal the customer table never sees; fed to the frozen TFM to test whether its
    hidden state adds anything over the raw factors."""
    from sklearn.decomposition import TruncatedSVD
    X = P.astype(np.float32).copy()
    X.data = np.log1p(X.data)
    n_components = int(min(n_components, X.shape[1] - 1, X.shape[0] - 1))
    Z = TruncatedSVD(n_components=n_components, random_state=seed).fit_transform(X)
    return standardize(Z.astype(np.float32))


def knn_cf(E: np.ndarray, P: sp.csr_matrix, query_rows: np.ndarray, k_neighbors: int = 50, K: int = 12,
           pop: np.ndarray | None = None, exclude_seen: bool = False, chunk: int = 1024,
           time_decay_weights: np.ndarray | None = None) -> np.ndarray:
    """For each query row: top-K article columns by cosine-weighted sum of neighbours' purchase rows."""
    En = E / np.clip(np.linalg.norm(E, axis=1, keepdims=True), 1e-9, None)
    En = En.astype(np.float32)
    n_items = P.shape[1]
    pop = np.zeros(n_items, np.float32) if pop is None else pop.astype(np.float32)
    out = np.zeros((len(query_rows), K), np.int64)
    for s in range(0, len(query_rows), chunk):
        q = query_rows[s:s + chunk]
        S = En[q] @ En.T                               # (chunk, n_customers)
        S[np.arange(len(q)), q] = -np.inf              # not yourself
        nb = np.argpartition(-S, k_neighbors, axis=1)[:, :k_neighbors]
        w = np.take_along_axis(S, nb, 1)
        w = np.clip(w, 0, None)
        scores = np.zeros((len(q), n_items), np.float32)
        for i in range(len(q)):
            scores[i] = w[i] @ P[nb[i]].toarray()
        scores += 1e-6 * pop
        if exclude_seen:
            seen = P[q].toarray() > 0
            scores[seen] = -np.inf
        top = np.argpartition(-scores, K, axis=1)[:, :K]
        order = np.argsort(-np.take_along_axis(scores, top, 1), axis=1)
        out[s:s + chunk] = np.take_along_axis(top, order, 1)
    return out


def map_at_k(pred: np.ndarray, truth_lists, K: int) -> float:
    """Mean average precision @K with per-row truth given as array-likes of destination ids (RelBench's formula)."""
    aps = []
    for p, t in zip(pred, truth_lists):
        t = set(np.atleast_1d(t).tolist())
        if not t:
            continue
        hits, score = 0, 0.0
        for i, a in enumerate(p[:K]):
            if a in t:
                hits += 1
                score += hits / (i + 1)
        aps.append(score / min(len(t), K))
    return float(np.mean(aps)) if aps else float("nan")


def knn_cf_sparse(E: sp.csr_matrix, P: sp.csr_matrix, query_rows: np.ndarray, k_neighbors: int = 50, K: int = 12,
                  pop: np.ndarray | None = None, chunk: int = 1024) -> np.ndarray:
    """knn_cf for a SPARSE embedding (e.g. the purchase matrix itself): cosine neighbours via sparse products."""
    norms = np.sqrt(np.asarray(E.multiply(E).sum(1)).ravel()) + 1e-9
    En = sp.diags(1 / norms) @ E
    EnT = En.T.tocsc()
    n_items = P.shape[1]
    pop = np.zeros(n_items, np.float32) if pop is None else pop.astype(np.float32)
    out = np.zeros((len(query_rows), K), np.int64)
    for s in range(0, len(query_rows), chunk):
        q = query_rows[s:s + chunk]
        S = (En[q] @ EnT).toarray().astype(np.float32)
        S[np.arange(len(q)), q] = -np.inf
        nb = np.argpartition(-S, k_neighbors, axis=1)[:, :k_neighbors]
        w = np.clip(np.take_along_axis(S, nb, 1), 0, None)
        scores = np.zeros((len(q), n_items), np.float32)
        for i in range(len(q)):
            scores[i] = w[i] @ P[nb[i]].toarray()
        scores += 1e-6 * pop
        top = np.argpartition(-scores, K, axis=1)[:, :K]
        order = np.argsort(-np.take_along_axis(scores, top, 1), axis=1)
        out[s:s + chunk] = np.take_along_axis(top, order, 1)
    return out


def past_visit(tx: pd.DataFrame, queries: np.ndarray, pop_articles: np.ndarray, K: int = 12,
               pad: bool = True) -> np.ndarray:
    """Most recent distinct past articles; padded with global popularity (pad=True) or -1 (pad=False)."""
    t = tx[tx["customer_id"].isin(queries)].sort_values("t_dat", ascending=False)
    last = t.groupby("customer_id")["article_id"].apply(lambda s: list(dict.fromkeys(s))[:K])
    filler = list(pop_articles) if pad else [-1] * K
    return np.array([(last.get(u, []) + filler)[:K] for u in queries])


def item_knn(P: sp.csr_matrix, query_rows: np.ndarray, K: int = 12, k_items: int = 50, pop: np.ndarray | None = None,
             chunk: int = 2048) -> np.ndarray:
    """Item-based CF: score(item) = sum over the user's past items of cosine(item, past item), top-K."""
    Pc = P.tocsc().astype(np.float32)
    norms = np.sqrt(np.asarray(Pc.multiply(Pc).sum(0)).ravel()) + 1e-9
    Pn = Pc.multiply(1 / norms).tocsc()
    S = (Pn.T @ Pn).tocsr()  # item x item cosine (sparse)
    pop = np.zeros(P.shape[1], np.float32) if pop is None else pop.astype(np.float32)
    out = np.zeros((len(query_rows), K), np.int64)
    for s in range(0, len(query_rows), chunk):
        q = query_rows[s:s + chunk]
        scores = (P[q] @ S).toarray() + 1e-6 * pop
        top = np.argpartition(-scores, K, axis=1)[:, :K]
        order = np.argsort(-np.take_along_axis(scores, top, 1), axis=1)
        out[s:s + chunk] = np.take_along_axis(top, order, 1)
    return out


def hybrid_fill(primary: np.ndarray, secondary: np.ndarray, K: int = 12) -> np.ndarray:
    """Primary list first (e.g. past purchases, -1 = empty slot), then secondary entries not already present."""
    out = np.zeros((len(primary), K), secondary.dtype)
    for i in range(len(primary)):
        seen, merged = set(), []
        for a in list(primary[i]) + list(secondary[i]):
            if a == -1 or a in seen:
                continue
            seen.add(a); merged.append(a)
            if len(merged) == K:
                break
        while len(merged) < K:
            merged.append(secondary[i][0])
        out[i] = merged
    return out


def run_split(task, db, split: str, embedders: list[str], hist_days: int = 365, k_neighbors: int = 50, seed: int = 0,
              device: str = "cuda", log: Any = None, sample_queries: int | None = None) -> dict[str, Any]:
    K = int(getattr(task, "eval_k", 12))
    tbl, cutoff, tx, queries = split_context(task, db, split, hist_days)
    if sample_queries and len(queries) > sample_queries:
        queries = np.random.default_rng(seed).choice(queries, sample_queries, replace=False)
        tbl = tbl[tbl["customer_id"].isin(queries)]
    art = db.table_dict["article"].df
    cust = db.table_dict["customer"].df
    articles = art["article_id"].to_numpy()
    pop_counts = tx["article_id"].value_counts()
    pop = pop_counts.reindex(articles).fillna(0).to_numpy(np.float32)
    pop_top = pop_counts.index[:K].to_numpy()
    subset = sample_queries is not None and len(tbl) < len(task.get_table(split).df)
    res: dict[str, Any] = {"split": split, "cutoff": str(cutoff.date()), "n_queries": len(queries), "hist_days": hist_days,
                           "K": K, "rows": {}, "official_evaluator": not subset}

    def evaluate(name: str, pred_articles: np.ndarray):
        m = pd.Series(np.arange(len(queries)), index=queries)
        aligned = pred_articles[m[tbl["customer_id"].to_numpy()].to_numpy()]  # one row per task-table row
        if subset:  # the official evaluator needs the full table; use the same MAP@K formula on the subset
            r = {"map": map_at_k(aligned, tbl["article_id"].to_numpy(), K)}
        elif split == "test":
            r = task.evaluate(aligned)  # relbench reads the hidden test targets itself
        else:
            r = task.evaluate(aligned, task.get_table(split))
        res["rows"][name] = {k: float(v) for k, v in r.items()}
        if log is not None:
            log.event("relbench_row", split=split, name=name, **res["rows"][name])
        print(f"[{split}] {name:28s} " + " ".join(f"{k}={100 * float(v):.3f}" for k, v in r.items()), flush=True)

    t0 = time.time()
    evaluate("GlobalPopularity", np.tile(pop_top, (len(queries), 1)))
    pv = past_visit(tx, queries, pop_top, K)
    pv_only = past_visit(tx, queries, pop_top, K, pad=False)
    evaluate("PastVisit", pv)
    P = purchase_matrix(tx, queries, articles)
    row, agg = customer_table(tx, cust, art, queries, cutoff)
    feats = {"row": standardize(row.to_numpy()), "agg": standardize(agg.to_numpy())}
    qrows = np.arange(len(queries))
    # reference rows that need no embedding model: classic collaborative filtering on the purchase matrix itself
    try:
        t1 = time.time()
        Pd = P.astype(np.float32)
        E_cf = Pd  # user-kNN over the raw purchase vectors (sparse cosine)
        pred = knn_cf_sparse(E_cf, P, qrows, k_neighbors=k_neighbors, K=K, pop=pop)
        evaluate("kNN-CF[purchase_matrix]", articles[pred])
        evaluate("Past+kNN-CF[purchase_matrix]", hybrid_fill(pv_only, articles[pred], K))
        res["rows"]["kNN-CF[purchase_matrix]"]["seconds"] = round(time.time() - t1, 1)
        t1 = time.time()
        pred = item_knn(P, qrows, K=K, pop=pop)
        evaluate("ItemKNN", articles[pred])
        evaluate("Past+ItemKNN", hybrid_fill(pv_only, articles[pred], K))
        res["rows"]["ItemKNN"]["seconds"] = round(time.time() - t1, 1)
    except Exception as e:
        print(f"[{split}] reference CF rows FAILED: {type(e).__name__}: {e}", flush=True)
    for name in embedders:
        t1 = time.time()
        try:
            if name in ("svd", "svd_agg", "tabpfn_svd_kmeans", "tabpfn_svd_random") and "svd" not in feats:
                feats["svd"] = svd_features(P, seed=seed)
                feats["svd_agg"] = np.hstack([feats["svd"], feats["agg"]])
            if name in feats:
                E = feats[name]
            elif name.startswith("kumo_relational"):
                from fer.relbench_kumo import embed_kumo_relational_relbench

                E = embed_kumo_relational_relbench(cust, art, tx, queries, agg=feats["agg"],
                                                   target="random" if name.endswith("random") else "kmeans", seed=seed, log=log)
            elif name == "tabpfn_svd_kmeans":
                E = embed_tabpfn(feats["svd_agg"], "kmeans", seed, device=device)
            elif name == "tabpfn_svd_random":
                E = embed_tabpfn(feats["svd_agg"], "random", seed, device=device)
            elif name == "tabpfn_agg_kmeans":
                E = embed_tabpfn(feats["agg"], "kmeans", seed, device=device)
            elif name == "tabpfn_agg_random":
                E = embed_tabpfn(feats["agg"], "random", seed, device=device)
            elif name == "tabpfn_row_kmeans":
                E = embed_tabpfn(feats["row"], "kmeans", seed, device=device)
            else:
                raise KeyError(name)
            pred = knn_cf(E, P, qrows, k_neighbors=k_neighbors, K=K, pop=pop)
            evaluate(f"kNN-CF[{name}]", articles[pred])
            evaluate(f"Past+kNN-CF[{name}]", hybrid_fill(pv_only, articles[pred], K))
            res["rows"][f"kNN-CF[{name}]"]["seconds"] = round(time.time() - t1, 1)
        except Exception as e:
            print(f"[{split}] {name} FAILED: {type(e).__name__}: {e}", flush=True)
            res["rows"][f"kNN-CF[{name}]"] = {"error": f"{type(e).__name__}: {e}"}
    res["seconds"] = round(time.time() - t0, 1)
    return res
