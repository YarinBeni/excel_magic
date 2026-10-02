"""Entity (customer) embedders compared on the retrieval benchmark.

row       : customer-table columns only (no relational information)
agg       : hand-crafted relational aggregates (RDL feature-engineering baseline)
gnn       : HeteroGraphSAGE trained from scratch with a customer-product link-prediction objective;
            the customer hidden states are the embedding
openrfm   : hidden state (256-d, pre-readout) of the pre-trained OpenRFM relational foundation model,
            fed the customer row + its last order lines + tickets as root/child/aux tables
openrfm_random : same architecture with random weights (control: does pre-training matter?)
"""
from __future__ import annotations

import sys
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from .db import ShopDB

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT / "third_party") not in sys.path:
    sys.path.insert(0, str(ROOT / "third_party"))


def _standardize(M: np.ndarray) -> np.ndarray:
    M = np.asarray(M, dtype=np.float32)
    mu, sd = np.nanmean(M, 0), np.nanstd(M, 0)
    sd[sd == 0] = 1
    return np.nan_to_num((M - mu) / sd)


def _cat_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])]


def _customer_frame(db: ShopDB) -> pd.DataFrame:
    c = db.customers.copy()
    c["tenure_days"] = (db.cutoff - pd.to_datetime(c.signup_date)).dt.days.clip(lower=0)
    c = c.drop(columns=["signup_date"])
    return c


# -------------------------------------------------------------------------------------------------
def embed_row(db: ShopDB, **_) -> np.ndarray:
    c = _customer_frame(db).set_index("customer_id")
    X = pd.get_dummies(c, columns=_cat_cols(c), dtype=float)
    return _standardize(X.to_numpy())


def embed_agg(db: ShopDB, **_) -> np.ndarray:
    c = _customer_frame(db).set_index("customer_id")
    past = db.past_items()
    ids = c.index
    g = past.groupby("customer_id")
    agg = pd.DataFrame(index=ids)
    agg["n_items"] = g.size().reindex(ids).fillna(0)
    agg["n_orders"] = g.order_id.nunique().reindex(ids).fillna(0)
    agg["spend"] = (past.price * past.quantity).groupby(past.customer_id).sum().reindex(ids).fillna(0)
    agg["days_since_last"] = (db.cutoff - g.order_date.max()).dt.days.reindex(ids).fillna(9999)
    mix = pd.crosstab(past.customer_id, past.category, normalize="index").reindex(ids).fillna(0)
    mix.columns = [f"cat_{x}" for x in mix.columns]
    tk = db.tickets[pd.to_datetime(db.tickets.created_date) <= db.cutoff] if len(db.tickets) else db.tickets
    agg["n_tickets"] = tk.groupby("customer_id").size().reindex(ids).fillna(0) if len(tk) else 0
    X = pd.concat([pd.get_dummies(c, columns=_cat_cols(c), dtype=float), agg, mix], axis=1)
    return _standardize(X.to_numpy())


# -------------------------------------------------------------------------------------------------
def embed_gnn(db: ShopDB, dim: int = 64, epochs: int = 60, seed: int = 0, log: Any = None, **_) -> np.ndarray:
    import torch
    import torch.nn.functional as F
    from torch_geometric.data import HeteroData
    from torch_geometric.nn import HeteroConv, SAGEConv

    torch.manual_seed(seed)
    cust = db.customers.reset_index(drop=True)
    prod = db.products.reset_index(drop=True)
    cid = {c: i for i, c in enumerate(cust.customer_id)}
    pid = {p: i for i, p in enumerate(prod.product_id)}
    past = db.past_items()
    src = torch.tensor(past.customer_id.map(cid).to_numpy())
    dst = torch.tensor(past.product_id.map(pid).to_numpy())
    data = HeteroData()
    data["customer"].x = torch.tensor(embed_row(db), dtype=torch.float32)
    px = pd.get_dummies(prod[["category"]].astype(str), dtype=float)
    px["price"] = prod.price.to_numpy()
    data["product"].x = torch.tensor(_standardize(px.to_numpy()), dtype=torch.float32)
    data["customer", "buys", "product"].edge_index = torch.stack([src, dst])
    data["product", "rev_buys", "customer"].edge_index = torch.stack([dst, src])

    class Net(torch.nn.Module):
        def __init__(self, dc, dp, h):
            super().__init__()
            self.lin = torch.nn.ModuleDict({"customer": torch.nn.Linear(dc, h), "product": torch.nn.Linear(dp, h)})
            self.c1 = HeteroConv({("customer", "buys", "product"): SAGEConv(h, h),
                                  ("product", "rev_buys", "customer"): SAGEConv(h, h)}, aggr="mean")
            self.c2 = HeteroConv({("customer", "buys", "product"): SAGEConv(h, h),
                                  ("product", "rev_buys", "customer"): SAGEConv(h, h)}, aggr="mean")

        def forward(self, x_dict, ei):
            h = {k: F.relu(self.lin[k](v)) for k, v in x_dict.items()}
            h1 = {k: F.relu(v) for k, v in self.c1(h, ei).items()}
            h2 = self.c2(h1, ei)
            return {k: h1[k] + h2[k] for k in h2}

    net = Net(data["customer"].x.shape[1], data["product"].x.shape[1], dim)
    opt = torch.optim.Adam(net.parameters(), lr=0.01)
    ei = data.edge_index_dict
    n_e = src.numel()
    t0 = time.time()
    for ep in range(epochs):
        net.train()
        opt.zero_grad()
        h = net(data.x_dict, ei)
        perm = torch.randperm(n_e)[: min(n_e, 8192)]
        pos = (h["customer"][src[perm]] * h["product"][dst[perm]]).sum(-1)
        neg_dst = torch.randint(0, len(prod), (perm.numel(),))
        neg = (h["customer"][src[perm]] * h["product"][neg_dst]).sum(-1)
        loss = F.binary_cross_entropy_with_logits(torch.cat([pos, neg]),
                                                  torch.cat([torch.ones_like(pos), torch.zeros_like(neg)]))
        loss.backward()
        opt.step()
        if log is not None and (ep % 20 == 0 or ep == epochs - 1):
            log.event("gnn_epoch", epoch=ep, loss=float(loss), t=round(time.time() - t0, 1))
    net.eval()
    with torch.no_grad():
        h = net(data.x_dict, ei)["customer"].numpy()
    return h


# -------------------------------------------------------------------------------------------------
def _openrfm_tables(db: ShopDB, rows_per_child: int = 10, rows_per_aux: int = 5, root_cols: int = 8,
                    child_cols: int = 8, aux_cols: int = 6):
    """Build root/child/aux arrays following OpenRFM's RelationalDataset._examples_to_arrays."""
    cust = _customer_frame(db).reset_index(drop=True)
    n = len(cust)
    # root: numeric + ordinal-coded categoricals, standardized
    root_df = cust.drop(columns=["customer_id"]).copy()
    for col in _cat_cols(root_df):
        root_df[col] = pd.factorize(root_df[col])[0]
    R = _standardize(root_df.to_numpy())[:, :root_cols]
    root = np.zeros((n, root_cols), np.float32)
    root[:, : R.shape[1]] = R

    past = db.past_items().sort_values("order_date")
    t0 = pd.to_datetime(db.orders.order_date).min()
    max_time = max(float((db.cutoff - t0).days), 1.0)
    past = past.assign(t=(past.order_date - t0).dt.days.astype(float),
                       cat_code=pd.factorize(past.category)[0].astype(float),
                       paid=(past.status == "paid").astype(float))
    feat_cols = ["cat_code", "price", "quantity", "paid"]
    past[feat_cols] = _standardize(past[feat_cols].to_numpy())
    anchor = max_time
    child = np.zeros((n, rows_per_child, child_cols), np.float32)
    child_mask = np.zeros((n, rows_per_child), bool)
    groups = {k: v for k, v in past.groupby("customer_id")}
    for i, c in enumerate(cust.customer_id):
        g = groups.get(c)
        if g is None or len(g) == 0:
            continue
        g = g.tail(rows_per_child)
        base = g[feat_cols].to_numpy(np.float32)
        rt = g.t.to_numpy(float)
        rel = np.maximum(anchor - rt, 0)
        tf = np.stack([rt / max_time, rel / max_time, np.log1p(rel) / np.log1p(max_time)], 1)
        vals = np.concatenate([base, tf], 1)[:, :child_cols]
        child[i, : len(vals), : vals.shape[1]] = vals
        child_mask[i, : len(vals)] = True

    aux = np.zeros((n, rows_per_aux, aux_cols), np.float32)
    aux_mask = np.zeros((n, rows_per_aux), bool)
    tk = db.tickets.copy()
    if len(tk):
        tk["created"] = pd.to_datetime(tk.created_date)
        tk = tk[tk.created <= db.cutoff].sort_values("created")
        tk["t"] = (tk.created - t0).dt.days.astype(float)
        tk["sev"] = tk.severity.map({"low": 0, "medium": 1, "high": 2}).fillna(0).astype(float)
        tk["open"] = (tk.status == "open").astype(float)
        tk[["sev", "open"]] = _standardize(tk[["sev", "open"]].to_numpy())
        tg = {k: v for k, v in tk.groupby("customer_id")}
        for i, c in enumerate(cust.customer_id):
            g = tg.get(c)
            if g is None or len(g) == 0:
                continue
            g = g.tail(rows_per_aux)
            rt = g.t.to_numpy(float)
            rel = np.maximum(anchor - rt, 0)
            tf = np.stack([rt / max_time, rel / max_time, np.log1p(rel) / np.log1p(max_time)], 1)
            vals = np.concatenate([g[["sev", "open"]].to_numpy(np.float32), tf], 1)[:, :aux_cols]
            aux[i, : len(vals), : vals.shape[1]] = vals
            aux_mask[i, : len(vals)] = True
    table_mask = np.ones((n, 3), bool)
    table_mask[:, 2] = aux_mask.any(1) | True  # keep aux table "available" (model pads empties)
    time_feat = np.tile(np.array([[1.0, 1.0]], np.float32), (n, 1))  # anchor at cutoff, local flag
    return root, child, aux, child_mask, aux_mask, table_mask, time_feat


def embed_openrfm(db: ShopDB, checkpoint: str | Path = ROOT / "weights/openrfm/openrfm-pretrain-ctx96.pt",
                  randomize: bool = False, batch_size: int = 64, context_size: int = 0, seed: int = 0,
                  log: Any = None, **_) -> np.ndarray:
    """Pre-readout hidden state of OpenRFM for every customer as the query example.

    context_size=0 feeds each customer alone (samples=1); context_size>0 additionally places K random
    other customers (targets hidden) in the in-context window so cross-sample attention is active.
    """
    import torch
    from kumorfm_repro.icl_eval import load_model_from_checkpoint

    torch.manual_seed(seed)
    model = load_model_from_checkpoint(Path(checkpoint), torch.device("cpu"))
    if randomize:
        for p in model.parameters():
            if p.dim() > 1:
                torch.nn.init.xavier_uniform_(p)
            else:
                torch.nn.init.zeros_(p)
    model.eval()
    lag_steps = model.root_encoder.col_id.shape[0] - 8 - 4  # task cols = 2 + lag + 2
    root, child, aux, cm, am, tm, tf = _openrfm_tables(db)
    n = len(root)
    rng = np.random.default_rng(seed)
    out = np.zeros((n, model.readout[0].normalized_shape[0]), np.float32)
    t0 = time.time()
    with torch.no_grad():
        for s in range(0, n, batch_size):
            idx = np.arange(s, min(n, s + batch_size))
            b = len(idx)
            S = context_size + 1
            # samples axis: [context..., query]; context examples are random other customers
            sel = np.stack([np.append(rng.choice(n, context_size, replace=False), i) for i in idx]) if context_size \
                else idx[:, None]
            def T(a, dtype=torch.float32, sel=sel):
                return torch.tensor(a[sel], dtype=dtype)
            qm = torch.zeros((b, S), dtype=torch.bool)
            qm[:, -1] = True
            ctx_t = torch.zeros((b, S))
            lag = torch.zeros((b, S, lag_steps))
            emb = model.encode(T(root), T(child), T(aux), T(cm, torch.bool), T(am, torch.bool), T(tm, torch.bool),
                               ctx_t, lag, T(tf), qm)
            out[idx] = emb.numpy()
            if log is not None and s == 0:
                log.event("openrfm_first_batch", seconds=round(time.time() - t0, 2), batch=b)
    if log is not None:
        log.event("openrfm_done", seconds=round(time.time() - t0, 1), n=n, dim=out.shape[1])
    return out


EMBEDDERS: dict[str, Callable[..., np.ndarray]] = {
    "row": embed_row,
    "agg": embed_agg,
    "gnn": embed_gnn,
    "openrfm": embed_openrfm,
    "openrfm_ctx16": lambda db, **kw: embed_openrfm(db, context_size=16, **kw),
    "openrfm_random": lambda db, **kw: embed_openrfm(db, randomize=True, **kw),
}


# -------------------------------------------------------------------------------------------------
def embed_kumo_relational(db: ShopDB, batch_size: int = 256, seed: int = 0, log: Any = None, **_) -> np.ndarray:
    """Graph embedding of each customer from NVIDIA's KumoRelational (sdm): the readout-token state that
    enters the ICL head, captured with a forward pre-hook (same pattern as sdm/examples/tabular/quickstart.py).

    Needs `tabfm-models download kumo-relational` (huggingface.co). UNTESTED until weights are available:
    written against the sdm 0.1 docstring API (RelatedTables / task_links / num_hops).
    """
    import torch
    from sdm import RelatedTables, TableTensor
    from sdm.models import KumoRelational
    from tabfm_auto.models import weights as W

    if not W.available("kumo-relational"):
        raise RuntimeError("kumo-relational weights not available: tabfm-models download kumo-relational")
    torch.manual_seed(seed)
    dev = "cpu"
    cust = db.customers.copy()
    cust["tenure_days"] = (db.cutoff - pd.to_datetime(cust.signup_date)).dt.days
    orders = db.orders[pd.to_datetime(db.orders.order_date) <= db.cutoff].copy()
    orders["order_date"] = pd.to_datetime(orders.order_date)
    items = db.items[db.items.order_id.isin(orders.order_id)].merge(db.products, on="product_id")
    items["category"] = pd.factorize(items.category)[0]
    tickets = db.tickets.copy()
    if len(tickets):
        tickets["created_date"] = pd.to_datetime(tickets.created_date)
        tickets = tickets[tickets.created_date <= db.cutoff]
        tickets["severity"] = tickets.severity.map({"low": 0, "medium": 1, "high": 2})
        tickets["open"] = (tickets.status == "open").astype(int)
        tickets = tickets.drop(columns=["status"])

    def tt(df, stypes):
        return TableTensor.from_pandas(df=df, stypes=stypes, device=dev)

    tables = {
        "customers": tt(cust[["customer_id", "age", "tenure_days"]], {"customer_id": "id", "age": "numerical", "tenure_days": "numerical"}),
        "orders": tt(orders[["order_id", "customer_id", "order_date"]], {"order_id": "id", "customer_id": "id", "order_date": "datetime"}),
        "order_items": tt(items[["item_id", "order_id", "category", "price", "quantity"]],
                          {"item_id": "id", "order_id": "id", "category": "numerical", "price": "numerical", "quantity": "numerical"}),
    }
    rels = [{"left_table": "orders", "left_columns": "customer_id", "right_table": "customers", "right_columns": "customer_id"},
            {"left_table": "order_items", "left_columns": "order_id", "right_table": "orders", "right_columns": "order_id"}]
    if len(tickets):
        tables["support_tickets"] = tt(tickets[["ticket_id", "customer_id", "created_date", "severity", "open"]],
                                       {"ticket_id": "id", "customer_id": "id", "created_date": "datetime",
                                        "severity": "numerical", "open": "numerical"})
        rels.append({"left_table": "support_tickets", "left_columns": "customer_id", "right_table": "customers",
                     "right_columns": "customer_id"})
    related = RelatedTables(tables=tables, relationships=rels,
                            task_links=[{"task_columns": "customer_id", "table": "customers", "table_columns": "customer_id"}])
    # a dummy binary target: the model needs *some* labelled context; targets are random so they carry no signal
    rng = np.random.default_rng(seed)
    task_df = pd.DataFrame({"customer_id": cust.customer_id.to_numpy(), "y": rng.integers(0, 2, len(cust)).astype(str)})
    task = tt(task_df, {"customer_id": "id", "y": "categorical"})
    model = KumoRelational(task="classification", device=dev)
    model.eval()
    captured: list[torch.Tensor] = []
    icl = model.models["classification"]
    head = getattr(getattr(icl, "icl_block", icl), "head", None)
    if head is None:
        raise RuntimeError("could not locate the ICL head to hook; inspect model.models['classification']")
    handle = head.register_forward_pre_hook(lambda m, args: captured.append(args[0].detach().float().cpu()))
    n = len(cust)
    n_ctx = min(64, n // 2)
    ctx_idx = rng.choice(n, n_ctx, replace=False)
    out = np.zeros((n, 0), np.float32)
    chunks = []
    with torch.no_grad():
        for s in range(0, n, batch_size):
            q_idx = np.arange(s, min(n, s + batch_size))
            captured.clear()
            model(x_context=task[ctx_idx].drop_columns("y"), y_context=task[ctx_idx, "y"],
                  x_query=task[q_idx].drop_columns("y"), related_context_tables=related, related_query_tables=related,
                  num_hops=2)
            h = captured[-1]
            h = h.reshape(-1, h.shape[-1])[-len(q_idx):]  # query rows are last along the sample axis
            chunks.append(h.numpy())
    handle.remove()
    out = np.concatenate(chunks, 0)
    if log is not None:
        log.event("kumo_relational_done", n=n, dim=int(out.shape[1]))
    return out


EMBEDDERS["kumo_relational"] = embed_kumo_relational


# -------------------------------------------------------------------------------------------------
def embed_tabpfn(db: ShopDB, table: str = "agg", target: str = "none", n_estimators: int = 1, seed: int = 0,
                 log: Any = None, **_) -> np.ndarray:
    """Hidden-state embeddings of the frozen TABULAR foundation model (TabPFN v2) for each customer row.

    ``table``  : which flat feature table to feed: "row" (customer columns only) or "agg" (hand relational aggregates).
    ``target`` : "none" fits TabPFN on a random binary target (no label signal; the embedding reflects the feature
                 structure only) or "churn" uses the post-cutoff churn label as a supervised anchor.
    The embedding is the per-row transformer state TabPFN exposes via ``get_embeddings`` (test-side tokens,
    averaged over estimators).
    """
    from tabfm_auto.models import get_model

    X = embed_row(db) if table == "row" else embed_agg(db)
    Xdf = pd.DataFrame(X, columns=[f"f{i}" for i in range(X.shape[1])])
    rng = np.random.default_rng(seed)
    if target == "churn" and db.hidden is not None:
        y = db.hidden.set_index("customer_id").churned.reindex(db.customer_ids()).to_numpy()
    else:
        y = rng.integers(0, 2, len(Xdf))
    n = len(Xdf)
    ctx = rng.choice(n, min(n, 1000), replace=False)  # TabPFN context (CPU budget); all rows are embedded as test rows
    clf = get_model(f"tabpfn:n_estimators={n_estimators}", "binary")
    t0 = time.time()
    clf.fit(Xdf.iloc[ctx], y[ctx])
    E = clf.get_embeddings(Xdf, data_source="test")
    E = np.asarray(E, dtype=np.float32)
    if E.ndim == 3:  # (n_estimators, n_rows, dim)
        E = E.mean(0)
    if log is not None:
        log.event("tabpfn_embed_done", table=table, target=target, seconds=round(time.time() - t0, 1), dim=int(E.shape[1]))
    return E


EMBEDDERS["tabpfn_row"] = lambda db, **kw: embed_tabpfn(db, table="row", **kw)
EMBEDDERS["tabpfn_agg"] = lambda db, **kw: embed_tabpfn(db, table="agg", **kw)
EMBEDDERS["tabpfn_agg_churn"] = lambda db, **kw: embed_tabpfn(db, table="agg", target="churn", **kw)


# -------------------------------------------------------------------------------------------------
# Target choice for the frozen tabular FM: the in-context label decides which directions the hidden state keeps.
def embed_tabpfn_multi(db: ShopDB, table: str = "agg", target: str = "random", n_targets: int = 4, seed: int = 0,
                       n_estimators: int = 1, log: Any = None, **_) -> np.ndarray:
    """Average TabPFN hidden states over several in-context targets (variance reduction for the random target).

    target: "random"  - n_targets independent random binary labels, embeddings averaged
            "kmeans"  - pseudo-labels from KMeans(k=8) on the (standardized) feature table, one fit
            "feature" - each of n_targets held-out feature columns (highest variance) used as a regression-style
                        binary target (above/below median) while the column itself is removed from the input
    """
    from sklearn.cluster import KMeans
    from tabfm_auto.models import get_model

    X = embed_row(db) if table == "row" else embed_agg(db)
    rng = np.random.default_rng(seed)
    n = len(X)
    embs = []

    def _embed(Xin: np.ndarray, y: np.ndarray) -> np.ndarray:
        Xdf = pd.DataFrame(Xin, columns=[f"f{i}" for i in range(Xin.shape[1])])
        ctx = rng.choice(n, min(n, 1000), replace=False)
        if len(np.unique(y[ctx])) < 2:
            y = y.copy(); y[ctx[0]] = 1 - y[ctx[0]]
        clf = get_model(f"tabpfn:n_estimators={n_estimators}", "binary")
        clf.fit(Xdf.iloc[ctx], y[ctx])
        E = np.asarray(clf.get_embeddings(Xdf, data_source="test"), dtype=np.float32)
        return E.mean(0) if E.ndim == 3 else E

    t0 = time.time()
    if target == "random":
        for _ in range(n_targets):
            embs.append(_embed(X, rng.integers(0, 2, n)))
    elif target == "kmeans":
        y = KMeans(8, n_init=4, random_state=seed).fit_predict(X)
        # TabPFN classifier handles multiclass; map to <=8 classes directly
        Xdf = pd.DataFrame(X, columns=[f"f{i}" for i in range(X.shape[1])])
        ctx = rng.choice(n, min(n, 1000), replace=False)
        clf = get_model(f"tabpfn:n_estimators={n_estimators}", "multiclass")
        clf.fit(Xdf.iloc[ctx], y[ctx])
        E = np.asarray(clf.get_embeddings(Xdf, data_source="test"), dtype=np.float32)
        embs.append(E.mean(0) if E.ndim == 3 else E)
    elif target == "feature":
        var = X.var(0)
        cols = np.argsort(-var)[:n_targets]
        for c in cols:
            y = (X[:, c] > np.median(X[:, c])).astype(int)
            Xin = np.delete(X, c, axis=1)
            embs.append(_embed(Xin, y))
    else:
        raise ValueError(target)
    E = np.mean(embs, axis=0)
    if log is not None:
        log.event("tabpfn_multi_done", table=table, target=target, n_targets=len(embs), seconds=round(time.time() - t0, 1))
    return E


EMBEDDERS["tabpfn_agg_rand4"] = lambda db, **kw: embed_tabpfn_multi(db, target="random", n_targets=4, **kw)
EMBEDDERS["tabpfn_agg_kmeans"] = lambda db, **kw: embed_tabpfn_multi(db, target="kmeans", **kw)
EMBEDDERS["tabpfn_agg_feat4"] = lambda db, **kw: embed_tabpfn_multi(db, target="feature", n_targets=4, **kw)
