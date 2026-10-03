"""Layer-wise row embeddings of a frozen *single-table* foundation model (TabPFN v2) on the same real RelBench entity
tasks as fer.relbench_layers, for a matched comparison with the relational model.

Each task row (entity, timestamp) becomes one flat feature row: the entity's own attributes plus, for every table that
references the entity table, the number of linked rows at or before the timestamp, the days since the latest one, and
the mean of each numeric column over the latest ``k`` linked rows (the same temporally safe window as the relational
subgraph). TabPFN is fitted on ``n_ctx`` training rows with one of four in-context targets:

  zeros   a constant label (what TEmBed's TabPFN approach uses for label-free row embeddings)
  random  random bits
  kmeans  k-means pseudo-labels on the standardised features (structure-preserving, label-free)
  label   the true training labels (the model's own in-context prediction is also scored)

and the label-column token of every query row is read after each of the 12 transformer blocks.
"""
from __future__ import annotations

import time
from typing import Any

import numpy as np
import pandas as pd

from fer.relbench_layers import Subgrapher, probe_scores


def row_features(sub: Subgrapher, entity_ids: np.ndarray, times: np.ndarray, k: int = 50) -> pd.DataFrame:
    rows = sub.ent_index.get_indexer(entity_ids) if sub.ent_index is not None else np.arange(len(entity_ids))
    out = sub.feat[sub.entity_table].iloc[np.clip(rows, 0, None)].reset_index(drop=True).add_prefix("ent_")
    out[rows < 0] = -1.0
    t_ns = pd.to_datetime(times).to_numpy()
    for c in sub.children:
        cs, groups, f = sub.s[c], sub.child_groups[c], sub.feat[c]
        tcol = cs["time"]
        tvals = pd.to_datetime(cs["df"][tcol]).to_numpy() if tcol else None
        num = f.to_numpy(np.float64) if len(f.columns) else np.zeros((len(cs["df"]), 0))
        pos = pd.Index(cs["df"].index)
        cnt = np.zeros(len(entity_ids)); since = np.full(len(entity_ids), -1.0)
        means = np.full((len(entity_ids), num.shape[1]), np.nan)
        for i, (e, t) in enumerate(zip(entity_ids, t_ns)):
            idx = groups.get(e)
            if idx is None or not len(idx):
                continue
            loc = pos.get_indexer(idx)
            if tcol:
                keep = tvals[loc] <= t
                loc = loc[keep]
            if not len(loc):
                continue
            cnt[i] = len(loc)
            if tcol:
                since[i] = (t - tvals[loc[-1]]) / np.timedelta64(1, "D")
            if num.shape[1]:
                means[i] = num[loc[-k:]].mean(0)
        out[f"{c}__count"] = cnt
        out[f"{c}__days_since_last"] = since
        for j, col in enumerate(f.columns):
            out[f"{c}__mean_{col}"] = means[:, j]
    return out.fillna(-1.0).astype(np.float32)


class BlockTap:
    """Hooks on TabPFN v2's transformer blocks; keeps the label-column token of the last ``n`` rows after each block."""

    def __init__(self, net):
        self.calls: dict[str, list] = {}
        self.handles = [b.register_forward_hook(self._hook(f"block_{i:02d}")) for i, b in enumerate(net.blocks)]

    def _hook(self, name):
        def hook(mod, args, out):
            o = out[0] if isinstance(out, tuple) else out
            self.calls.setdefault(name, []).append(o.detach().float().cpu())
        return hook

    def take(self, n: int) -> dict[str, np.ndarray]:
        res = {}
        for name, lst in self.calls.items():
            x = lst[-1]  # [B, R, C, D]
            res[name] = x.reshape(-1, *x.shape[-3:])[0, -n:, -1, :].numpy()
        self.calls = {}
        return res

    def remove(self):
        for h in self.handles:
            h.remove()


def tabpfn_layer_embeddings(X_ctx: pd.DataFrame, y_ctx: np.ndarray, X_sets: list[pd.DataFrame], device: str = "cuda",
                            chunk: int = 2000) -> tuple[list[dict[str, np.ndarray]], list[np.ndarray]]:
    from tabfm_auto.models import get_model

    n_cls = len(np.unique(y_ctx))
    clf = get_model(f"tabpfn:n_estimators=1,device={device},inference_precision=float32",
                    "binary" if n_cls <= 2 else "multiclass")
    clf.fit(X_ctx, y_ctx)
    tap = BlockTap(clf.models_[0])
    outs, scores = [], []
    try:
        for X in X_sets:
            parts: dict[str, list] = {}
            sc = []
            for s in range(0, len(X), chunk):
                xc = X.iloc[s:s + chunk]
                final = clf.get_embeddings(xc, data_source="test")
                got = tap.take(len(xc))
                final = np.asarray(final, np.float32).reshape(-1, len(xc), np.asarray(final).shape[-1])[0]
                got["final_api"] = final
                for k2, v in got.items():
                    parts.setdefault(k2, []).append(v)
                if n_cls == 2 and set(np.unique(y_ctx)) <= {0, 1}:
                    sc.append(clf.predict_proba(xc)[:, 1])
            outs.append({k2: np.concatenate(v, 0) for k2, v in parts.items()})
            scores.append(np.concatenate(sc) if sc else None)
    finally:
        tap.remove()
    return outs, scores


def run_task(dataset: str, task_name: str, targets=("zeros", "random", "kmeans", "label"), n_ctx: int = 3000,
             n_train: int = 4000, k_children: int = 50, seed: int = 0, device: str = "cuda", log: Any = None,
             max_eval: int | None = 20000) -> dict[str, Any]:
    import relbench
    from sklearn.cluster import KMeans
    from sklearn.metrics import roc_auc_score
    from sklearn.preprocessing import StandardScaler

    ds = relbench.load_dataset(dataset)
    task = ds.load_task(task_name)
    db = task.get_db(upto_test_timestamp=False)
    ent_col, time_col, tgt = task.entity_col, task.time_col, task.target_col
    tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
    va = task.get_table("val", mask_input_cols=False).df.reset_index(drop=True)
    te = task.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
    rng = np.random.default_rng(seed)
    if max_eval and len(va) > max_eval:
        va = va.iloc[np.sort(rng.choice(len(va), max_eval, replace=False))].reset_index(drop=True)
    sub = Subgrapher(db, task.entity_table, k_children=k_children)
    perm = rng.permutation(len(tr))
    ctx = tr.iloc[perm[:min(n_ctx, len(tr) // 2)]].reset_index(drop=True)
    prb = tr.iloc[perm[min(n_ctx, len(tr) // 2):][:n_train]].reset_index(drop=True)
    t0 = time.time()
    F = {name: row_features(sub, d[ent_col].to_numpy(), d[time_col].to_numpy(), k_children)
         for name, d in (("ctx", ctx), ("prb", prb), ("va", va), ("te", te))}
    res: dict[str, Any] = {"dataset": dataset, "task": task_name, "n_features": int(F["ctx"].shape[1]),
                           "feature_seconds": round(time.time() - t0, 1), "n_ctx": len(ctx), "n_probe_train": len(prb),
                           "n_val": len(va), "n_test": len(te), "targets": {}}
    for target in targets:
        if target == "zeros":
            y_ctx = np.zeros(len(ctx), int)
        elif target == "random":
            y_ctx = rng.integers(0, 2, len(ctx))
        elif target == "kmeans":
            y_ctx = KMeans(8, n_init=4, random_state=seed).fit_predict(StandardScaler().fit_transform(F["ctx"]))
        else:
            y_ctx = ctx[tgt].to_numpy().astype(int)
        m: dict[str, Any] = {"layers": {}}
        try:
            (Etr, Eva, Ete), (_, s_va, s_te) = tabpfn_layer_embeddings(F["ctx"], y_ctx, [F["prb"], F["va"], F["te"]], device)
        except Exception as e:  # e.g. a constant target refused by the model
            m["error"] = f"{type(e).__name__}: {str(e)[:300]}"
            res["targets"][target] = m
            continue
        if target == "label" and s_te is not None:
            m["icl_head"] = {"val_auroc": float(roc_auc_score(va[tgt], s_va)), "test": {k: float(v) for k, v in task.evaluate(s_te).items()}}
        ytr = prb[tgt].to_numpy().astype(int)
        m["final_api_matches_last_block"] = bool(np.allclose(Ete["final_api"], Ete[max(k for k in Ete if k.startswith("block_"))], atol=1e-3))
        layers = ["raw_features"] + sorted(k for k in Etr if k.startswith("block_"))
        for layer in layers:
            a, b, c = (F["prb"].to_numpy(), F["va"].to_numpy(), F["te"].to_numpy()) if layer == "raw_features" else (Etr[layer], Eva[layer], Ete[layer])
            pv, pt = probe_scores(a, ytr, b), probe_scores(a, ytr, c)
            m["layers"][layer] = {p: {"val_auroc": float(roc_auc_score(va[tgt], pv[p])),
                                      "test": {k: float(v) for k, v in task.evaluate(pt[p]).items()}} for p in pv}
            if log is not None:
                log.event("tabular_layer_probe", dataset=dataset, task=task_name, target=target, layer=layer,
                          **{f"{p}_test": m["layers"][layer][p]["test"].get("roc_auc") for p in pv})
        for p in ("linear", "knn"):
            cand = [L for L in m["layers"] if L != "raw_features"]
            best = max(cand, key=lambda L: m["layers"][L][p]["val_auroc"])
            m[f"best_by_val_{p}"] = {"layer": best, **m["layers"][best][p]}
        res["targets"][target] = m
    return res
