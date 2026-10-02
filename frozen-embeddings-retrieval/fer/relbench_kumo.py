"""Frozen Kumo Relational (NVIDIA sdm) customer embeddings from the RelBench rel-hm tables.

Same mechanism as fer.embedders.embed_kumo_relational on the synthetic shop DB: the readout-token state that enters the
in-context head, captured with a forward pre-hook; the in-context target is k-means pseudo-labels on the hand
aggregates (``target="kmeans"``) or random bits (``target="random"``). The relational graph is customers <- transactions
-> articles, restricted to the probed customers and the history window.
"""
from __future__ import annotations

import os
import time
from typing import Any

import numpy as np
import pandas as pd


def _num(s: pd.Series) -> pd.Series:
    if s.dtype == object or str(s.dtype).startswith("str") or str(s.dtype) == "category":
        return pd.Series(pd.factorize(s.astype(str))[0], index=s.index).astype(float)
    return pd.to_numeric(s, errors="coerce").fillna(-1).astype(float)


def embed_kumo_relational_relbench(cust: pd.DataFrame, art: pd.DataFrame, tx: pd.DataFrame, customers: np.ndarray,
                                   agg: np.ndarray | None = None, target: str = "kmeans", n_clusters: int = 8,
                                   seed: int = 0, batch_size: int = 256, n_context: int = 64, num_hops: int = 2,
                                   device: str | None = None, max_tx_per_customer: int = 200, log: Any = None) -> np.ndarray:
    import torch
    from sdm import RelatedTables, TableTensor
    from sdm.models import KumoRelational

    dev = device or os.environ.get("FER_DEVICE", "cuda" if torch.cuda.is_available() else "cpu")
    torch.manual_seed(seed)
    rng = np.random.default_rng(seed)
    cidx = pd.Index(customers)
    c = cust.set_index("customer_id").reindex(cidx).reset_index()
    cdf = pd.DataFrame({"customer_id": np.arange(len(cidx))})
    for col in ("age", "club_member_status", "fashion_news_frequency", "Active", "FN"):
        if col in c.columns:
            cdf[col] = _num(c[col]).to_numpy()
    t = tx[tx["customer_id"].isin(cidx)].copy()
    t = t.sort_values("t_dat").groupby("customer_id").tail(max_tx_per_customer)  # cap the fan-out
    t["customer_id"] = cidx.get_indexer(t["customer_id"])
    aidx = pd.Index(t["article_id"].unique())
    t["article_id"] = aidx.get_indexer(t["article_id"])
    tdf = pd.DataFrame({"tx_id": np.arange(len(t)), "customer_id": t["customer_id"].to_numpy(),
                        "article_id": t["article_id"].to_numpy(), "t_dat": pd.to_datetime(t["t_dat"]).to_numpy(),
                        "price": pd.to_numeric(t["price"], errors="coerce").fillna(0).to_numpy(float),
                        "sales_channel_id": pd.to_numeric(t["sales_channel_id"], errors="coerce").fillna(0).to_numpy(float)})
    a = art.set_index("article_id").reindex(aidx).reset_index()
    adf = pd.DataFrame({"article_id": np.arange(len(aidx))})
    for col in ("product_type_no", "graphical_appearance_no", "colour_group_code", "department_no", "index_group_no",
                "garment_group_no", "section_no"):
        if col in a.columns:
            adf[col] = _num(a[col]).to_numpy()

    def tt(df, stypes):
        return TableTensor.from_pandas(df=df, stypes=stypes, device=dev)

    tables = {
        "customers": tt(cdf, {"customer_id": "id", **{k: "numerical" for k in cdf.columns if k != "customer_id"}}),
        "transactions": tt(tdf, {"tx_id": "id", "customer_id": "id", "article_id": "id", "t_dat": "datetime",
                                 "price": "numerical", "sales_channel_id": "numerical"}),
        "articles": tt(adf, {"article_id": "id", **{k: "numerical" for k in adf.columns if k != "article_id"}}),
    }
    rels = [{"left_table": "transactions", "left_columns": "customer_id", "right_table": "customers", "right_columns": "customer_id"},
            {"left_table": "transactions", "left_columns": "article_id", "right_table": "articles", "right_columns": "article_id"}]
    related = RelatedTables(tables=tables, relationships=rels,
                            task_links=[{"task_columns": "customer_id", "table": "customers", "table_columns": "customer_id"}])
    n = len(cidx)
    if target == "kmeans":
        from sklearn.cluster import KMeans

        if agg is None:
            raise ValueError("target='kmeans' needs the aggregate features")
        y_ctx = KMeans(n_clusters, n_init=4, random_state=seed).fit_predict(agg)
    else:
        y_ctx = rng.integers(0, 2, n)
    task = tt(pd.DataFrame({"customer_id": np.arange(n), "y": y_ctx.astype(str)}), {"customer_id": "id", "y": "categorical"})
    model = KumoRelational(task="classification", device=dev)
    model.eval()
    captured: list[torch.Tensor] = []
    icl = model.models["classification"]
    head = getattr(getattr(icl, "icl_block", icl), "head", None)
    if head is None:
        raise RuntimeError("could not locate the ICL head to hook")
    handle = head.register_forward_pre_hook(lambda m, args: captured.append(args[0].detach().float().cpu()))
    ctx_idx = rng.choice(n, min(n_context, n // 2), replace=False)
    chunks = []
    t0 = time.time()
    with torch.no_grad():
        for s in range(0, n, batch_size):
            q_idx = np.arange(s, min(n, s + batch_size))
            captured.clear()
            model(x_context=task[ctx_idx].drop_columns("y"), y_context=task[ctx_idx, "y"],
                  x_query=task[q_idx].drop_columns("y"), related_context_tables=related, related_query_tables=related,
                  num_hops=num_hops)
            h = captured[-1]
            chunks.append(h.reshape(-1, h.shape[-1])[-len(q_idx):].numpy())
            if log is not None and (s // batch_size) % 20 == 0:
                log.event("kumo_relational_progress", done=int(min(n, s + batch_size)), n=n, seconds=round(time.time() - t0, 1))
    handle.remove()
    out = np.concatenate(chunks, 0).astype(np.float32)
    if log is not None:
        log.event("kumo_relational_done", n=n, dim=int(out.shape[1]), seconds=round(time.time() - t0, 1))
    return out
