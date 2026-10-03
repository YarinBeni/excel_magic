"""Layer-wise entity embeddings of a frozen relational foundation model on real RelBench entity tasks.

For a RelBench entity-classification task (official train / val / test tables, official test evaluator), embed every
task row with frozen Kumo Relational (NVIDIA sdm) and read the hidden state of the query row at every stage:

  pre_gnn   per-table row embedding of the entity row (no relational context yet)
  gnn       after relational message passing (graph-aware, before in-context learning)
  icl_00..  output of each in-context transformer layer (12)
  final     the normalised state that enters the prediction head

Two context modes: ``random`` (label-free: random bits as the in-context target, a task-agnostic embedding) and
``label`` (the true training labels in context: the model's own in-context prediction, comparable to the
leaderboard's in-context entries). Each layer is scored with a logistic-regression probe and a kNN probe trained on
train-split embeddings; the layer is chosen on validation and its test AUROC reported with the official evaluator.

Subgraph per task row (2 hops, temporally safe): the entity row; rows of every table with a foreign key to the entity
table, timestamped at or before the task row's timestamp, latest ``k_children`` per row; and the parent rows those
children reference. Each task row gets its own copy of its entity and children, so one forward can mix timestamps.
"""
from __future__ import annotations

import time
from typing import Any

import numpy as np
import pandas as pd

ID_MAX_UNIQUE_STR = 2000  # string columns with more distinct values than this are free text: dropped


# ----------------------------------------------------------------------------------------------- schema
def table_schema(db) -> dict[str, dict[str, Any]]:
    out = {}
    for name, t in db.table_dict.items():
        out[name] = {"pkey": t.pkey_col, "fkeys": dict(t.fkey_col_to_pkey_table), "time": t.time_col, "df": t.df}
    return out


def _numeric_frame(df: pd.DataFrame, drop: set[str]) -> pd.DataFrame:
    cols = {}
    for c in df.columns:
        if c in drop:
            continue
        s = df[c]
        if pd.api.types.is_bool_dtype(s):
            cols[c] = s.astype(float)
        elif pd.api.types.is_numeric_dtype(s):
            cols[c] = pd.to_numeric(s, errors="coerce").astype(float)
        elif pd.api.types.is_datetime64_any_dtype(s):
            cols[c] = (pd.to_datetime(s) - pd.Timestamp("2000-01-01")).dt.days.astype(float)
        else:
            nun = s.nunique(dropna=True)
            if 0 < nun <= ID_MAX_UNIQUE_STR:
                cols[c] = pd.Series(pd.factorize(s.astype(str))[0], index=s.index).astype(float)
    f = pd.DataFrame(cols, index=df.index)
    return f.fillna(-1.0) if len(f.columns) else f


class Subgrapher:
    """Builds sdm RelatedTables for a batch of task rows (entity id, timestamp)."""

    def __init__(self, db, entity_table: str, k_children: int = 50):
        self.s = table_schema(db)
        self.entity_table = entity_table
        self.k = k_children
        ent = self.s[entity_table]
        self.children = [n for n, t in self.s.items() if n != entity_table and entity_table in t["fkeys"].values()]
        self.child_fk = {n: next(c for c, p in self.s[n]["fkeys"].items() if p == entity_table) for n in self.children}
        # numeric feature frames (keys and the time column excluded; time enters as the datetime column)
        self.feat = {}
        for n, t in self.s.items():
            drop = {t["pkey"], *t["fkeys"].keys()} - {None}
            if t["time"]:
                drop.add(t["time"])
            self.feat[n] = _numeric_frame(t["df"], drop)
        self.ent_index = pd.Index(ent["df"][ent["pkey"]]) if ent["pkey"] else None
        # per child: rows grouped by entity, sorted by time
        self.child_groups = {}
        for n in self.children:
            df = self.s[n]["df"]
            tcol = self.s[n]["time"]
            order = df.sort_values(tcol).index if tcol else df.index
            self.child_groups[n] = df.loc[order].groupby(self.child_fk[n], sort=False).indices
            self.child_groups[n] = {k: df.loc[order].index.to_numpy()[v] for k, v in self.child_groups[n].items()}

    def build(self, entity_ids: np.ndarray, times: np.ndarray, device):
        from sdm import RelatedTables, TableTensor

        n = len(entity_ids)
        t_ns = pd.to_datetime(times)
        tables, rels = {}, []
        ent = self.s[self.entity_table]
        rows = self.ent_index.get_indexer(entity_ids) if self.ent_index is not None else np.arange(n)
        efeat = self.feat[self.entity_table].iloc[np.clip(rows, 0, None)].reset_index(drop=True)
        efeat[rows < 0] = -1.0
        edf = pd.DataFrame({"__eid": np.arange(n)})
        edf = pd.concat([edf, efeat], axis=1)
        if ent["time"]:
            edf["__t"] = pd.to_datetime(ent["df"][ent["time"]].iloc[np.clip(rows, 0, None)].to_numpy())
        tables[self.entity_table] = (edf, {"__eid": "id", **({"__t": "datetime"} if ent["time"] else {})})
        min_t = t_ns.min()
        for c in self.children:
            cs = self.s[c]
            groups = self.child_groups[c]
            tcol = cs["time"]
            take, owner = [], []
            for i, (e, t) in enumerate(zip(entity_ids, t_ns)):
                idx = groups.get(e)
                if idx is None or not len(idx):
                    continue
                if tcol:
                    ts = cs["df"].loc[idx, tcol].to_numpy()
                    idx = idx[ts <= np.datetime64(t)]
                idx = idx[-self.k:]
                take.append(idx); owner.append(np.full(len(idx), i))
            if not take:
                continue
            take = np.concatenate(take); owner = np.concatenate(owner)
            cdf = pd.DataFrame({"__eid": owner})
            cdf = pd.concat([cdf, self.feat[c].loc[take].reset_index(drop=True)], axis=1)
            st = {"__eid": "id"}
            if tcol:
                cdf["__t"] = pd.to_datetime(cs["df"].loc[take, tcol].to_numpy()); st["__t"] = "datetime"
            # parents referenced by these child rows (other than the entity table)
            for fk, ptab in cs["fkeys"].items():
                if ptab == self.entity_table or ptab not in self.s or self.s[ptab]["pkey"] is None:
                    continue
                ps = self.s[ptab]
                ref = cs["df"].loc[take, fk].to_numpy()
                pindex = pd.Index(ps["df"][ps["pkey"]])
                keep = pd.unique(ref[pd.notna(ref)])
                prow = pindex.get_indexer(keep)
                keep, prow = keep[prow >= 0], prow[prow >= 0]
                if ps["time"]:
                    pt = pd.to_datetime(ps["df"][ps["time"]].iloc[prow].to_numpy())
                    ok = np.asarray(pt <= min_t)
                    keep, prow = keep[ok], prow[ok]
                local = pd.Series(np.arange(len(keep)), index=keep)
                pname = f"{ptab}__via_{c}"  # one parent copy per child table keeps the ids local and consistent
                pdf = pd.concat([pd.DataFrame({"__pid": np.arange(len(keep))}),
                                 self.feat[ptab].iloc[prow].reset_index(drop=True)], axis=1)
                tables[pname] = (pdf, {"__pid": "id"})
                lk = local.reindex(ref).to_numpy()
                cdf[f"__fk_{fk}"] = np.where(np.isnan(lk), -1, lk).astype(np.int64)
                st[f"__fk_{fk}"] = "id"
                rels.append({"left_table": c, "left_columns": f"__fk_{fk}", "right_table": pname, "right_columns": "__pid"})
            tables[c] = (cdf, st)
            rels.append({"left_table": c, "left_columns": "__eid", "right_table": self.entity_table, "right_columns": "__eid"})
        tt = {}
        for name, (df, st) in tables.items():
            stp = {col: st.get(col, "numerical") for col in df.columns}
            tt[name] = TableTensor.from_pandas(df=df, stypes=stp, device=device)
        related = RelatedTables(tables=tt, relationships=rels,
                                task_links=[{"task_columns": "__eid", "table": self.entity_table, "table_columns": "__eid"}])
        task = TableTensor.from_pandas(df=pd.DataFrame({"__eid": np.arange(n), "__t": t_ns}),
                                       stypes={"__eid": "id", "__t": "datetime"}, device=device)
        return task, related


# ----------------------------------------------------------------------------------------------- hooks
class LayerTap:
    """Forward hooks on the frozen model that keep the query rows' hidden state at every stage."""

    def __init__(self, inner):
        self.calls: dict[str, list] = {}
        self.handles = []

        def pre_gnn(mod, args, kwargs):
            x = args[0] if len(args) > 0 else kwargs["x"]
            graph = args[1] if len(args) > 1 else kwargs["graph"]
            rt, ri = kwargs["readout_table"], kwargs["readout_index"]
            s, e = graph.start_node_offsets[rt], graph.end_node_offsets[rt]
            self.calls.setdefault("pre_gnn", []).append(x[s:e][ri].detach().float().cpu().clone())

        def post(name):
            def hook(mod, args, out):
                o = out[0] if isinstance(out, tuple) else out
                self.calls.setdefault(name, []).append(o.detach().float().cpu().clone())
            return hook

        self.handles.append(inner.gnn.register_forward_pre_hook(pre_gnn, with_kwargs=True))
        self.handles.append(inner.gnn.register_forward_hook(post("gnn")))
        for i, layer in enumerate(inner.icl_block.layers):
            self.handles.append(layer.register_forward_hook(post(f"icl_{i:02d}")))
        self.handles.append(inner.icl_block.norm.register_forward_hook(post("final")))

    def reset(self):
        self.calls = {}

    def query_states(self, n_query: int) -> dict[str, np.ndarray]:
        out = {}
        for name, lst in self.calls.items():
            h = lst[-1]  # the gnn runs for the context, then for the query; ICL layers run once
            h = h.reshape(-1, h.shape[-1]) if h.dim() == 2 else h.reshape(-1, h.shape[-2], h.shape[-1])[0]
            out[name] = h[-n_query:].numpy()
        return out

    def remove(self):
        for h in self.handles:
            h.remove()


# ----------------------------------------------------------------------------------------------- embedding
def _positive_scores(out, n: int) -> np.ndarray:
    """The model returns a TableTensor with one numerical column per class label (e.g. ('1', '0'))."""
    import torch

    key = next(k for k in out.columns if str(k).endswith("numerical"))
    names = [str(c) for c in out.columns[key]]
    probs = torch.as_tensor(out.numerical).float().cpu().numpy().reshape(-1, len(names))[-n:]
    pos = names.index("1") if "1" in names else (names.index("True") if "True" in names else 0)
    return probs[:, pos]


def embed_rows(model, sub: Subgrapher, ctx: pd.DataFrame, ctx_y: np.ndarray, rows: pd.DataFrame, ent_col: str,
               time_col: str, device, batch: int = 256, log: Any = None) -> tuple[dict[str, np.ndarray], np.ndarray]:
    """Hidden states (per layer) and the model's own positive-class score for ``rows``."""
    import torch
    from sdm import TableTensor

    inner = next(iter(model.models.values()))
    tap = LayerTap(inner)
    cx, crel = sub.build(ctx[ent_col].to_numpy(), ctx[time_col].to_numpy(), device)
    cy = TableTensor.from_pandas(df=pd.DataFrame({"y": ctx_y.astype(int).astype(str)}), stypes={"y": "categorical"},
                                 device=device)
    states: dict[str, list] = {}
    scores = []
    order = np.argsort(pd.to_datetime(rows[time_col]).to_numpy(), kind="stable")
    t0 = time.time()
    try:
        with torch.no_grad():
            for s in range(0, len(rows), batch):
                idx = order[s:s + batch]
                r = rows.iloc[idx]
                qx, qrel = sub.build(r[ent_col].to_numpy(), r[time_col].to_numpy(), device)
                tap.reset()
                torch.manual_seed(0)  # the GNN draws random edge-type embeddings: same draw for every batch
                out = model(x_context=cx, y_context=cy, x_query=qx, related_context_tables=crel,
                            related_query_tables=qrel, num_hops=2)
                for k, v in tap.query_states(len(idx)).items():
                    states.setdefault(k, []).append((idx, v))
                scores.append((idx, _positive_scores(out, len(idx))))
                if log is not None and (s // batch) % 20 == 0:
                    log.event("embed_progress", done=int(min(len(rows), s + batch)), n=len(rows),
                              seconds=round(time.time() - t0, 1))
    finally:
        tap.remove()
    full = {}
    for k, parts in states.items():
        a = np.zeros((len(rows), parts[0][1].shape[1]), np.float32)
        for idx, v in parts:
            a[idx] = v
        full[k] = a
    sc = np.zeros(len(rows), np.float32)
    for idx, v in scores:
        sc[idx] = v
    return full, sc


# ----------------------------------------------------------------------------------------------- probes
def probe_scores(E_tr, y_tr, E_ev, k: int = 50) -> dict[str, np.ndarray]:
    from sklearn.linear_model import LogisticRegression
    from sklearn.neighbors import NearestNeighbors
    from sklearn.preprocessing import StandardScaler

    sc = StandardScaler().fit(E_tr)
    A, B = sc.transform(E_tr), sc.transform(E_ev)
    lr = LogisticRegression(max_iter=3000, C=0.1).fit(A, y_tr)
    out = {"linear": lr.predict_proba(B)[:, 1]}
    An = A / (np.linalg.norm(A, axis=1, keepdims=True) + 1e-9)
    Bn = B / (np.linalg.norm(B, axis=1, keepdims=True) + 1e-9)
    nn = NearestNeighbors(n_neighbors=min(k, len(An)), metric="cosine").fit(An)
    _, ix = nn.kneighbors(Bn)
    out["knn"] = y_tr[ix].mean(1)
    return out


def run_task(dataset: str, task_name: str, modes=("random", "label"), n_ctx: int = 512, n_train: int = 4000,
             k_children: int = 50, seed: int = 0, device: str = "cuda", log: Any = None, max_eval: int | None = None) -> dict[str, Any]:
    import relbench
    import torch
    from sdm.models import KumoRelational
    from sklearn.metrics import roc_auc_score

    ds = relbench.load_dataset(dataset)
    task = ds.load_task(task_name)
    db = task.get_db(upto_test_timestamp=False)
    ent_col, ent_table, time_col, tgt = task.entity_col, task.entity_table, task.time_col, task.target_col
    tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
    va = task.get_table("val", mask_input_cols=False).df.reset_index(drop=True)
    te = task.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
    rng = np.random.default_rng(seed)
    if max_eval and len(va) > max_eval:
        va = va.iloc[np.sort(rng.choice(len(va), max_eval, replace=False))].reset_index(drop=True)
    sub = Subgrapher(db, ent_table, k_children=k_children)
    # context: unique entities, stratified over labels, latest timestamps first
    trs = tr.sample(frac=1.0, random_state=seed).drop_duplicates(ent_col)
    pos, neg = trs[trs[tgt] == 1], trs[trs[tgt] != 1]
    nc_pos = min(len(pos), max(n_ctx // 2, int(n_ctx * len(pos) / max(len(trs), 1))))
    ctx_all = pd.concat([pos.head(nc_pos), neg.head(n_ctx - nc_pos)])
    rest = tr.drop(index=ctx_all.index)
    ctx = ctx_all.reset_index(drop=True)
    prb = rest.iloc[rng.choice(len(rest), min(n_train, len(rest)), replace=False)].reset_index(drop=True)
    torch.manual_seed(seed)
    model = KumoRelational(device=device)
    model.eval()
    res: dict[str, Any] = {"dataset": dataset, "task": task_name, "n_ctx": len(ctx), "n_probe_train": len(prb),
                           "n_val": len(va), "n_test": len(te), "pos_rate_train": float(tr[tgt].mean()),
                           "tables": list(sub.s), "children": sub.children, "modes": {}}
    for mode in modes:
        y_ctx = ctx[tgt].to_numpy().astype(int) if mode == "label" else rng.integers(0, 2, len(ctx))
        t0 = time.time()
        Etr, s_tr = embed_rows(model, sub, ctx, y_ctx, prb, ent_col, time_col, device, log=log)
        Eva, s_va = embed_rows(model, sub, ctx, y_ctx, va, ent_col, time_col, device, log=log)
        Ete, s_te = embed_rows(model, sub, ctx, y_ctx, te, ent_col, time_col, device, log=log)
        m: dict[str, Any] = {"embed_seconds": round(time.time() - t0, 1), "layers": {}}
        if mode == "label":
            m["icl_head"] = {"val_auroc": float(roc_auc_score(va[tgt], s_va)), "test": {k: float(v) for k, v in task.evaluate(s_te).items()}}
        ytr = prb[tgt].to_numpy().astype(int)
        for layer in Etr:
            pv = probe_scores(Etr[layer], ytr, Eva[layer])
            pt = probe_scores(Etr[layer], ytr, Ete[layer])
            m["layers"][layer] = {p: {"val_auroc": float(roc_auc_score(va[tgt], pv[p])),
                                      "test": {k: float(v) for k, v in task.evaluate(pt[p]).items()}} for p in pv}
            if log is not None:
                log.event("layer_probe", dataset=dataset, task=task_name, mode=mode, layer=layer,
                          **{f"{p}_val": m["layers"][layer][p]["val_auroc"] for p in pv},
                          **{f"{p}_test": m["layers"][layer][p]["test"].get("roc_auc") for p in pv})
        for p in ("linear", "knn"):
            best = max(m["layers"], key=lambda L: m["layers"][L][p]["val_auroc"])
            m[f"best_by_val_{p}"] = {"layer": best, **m["layers"][best][p]}
        res["modes"][mode] = m
    return res
