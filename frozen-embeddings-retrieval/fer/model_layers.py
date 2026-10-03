"""Layer-wise row embeddings for every open tabular foundation model, plus label-free geometry of each layer.

Models (all frozen): TabPFN v2, TabPFN-2.5 (12 transformer blocks; the label-column token of each query row), and the
sdm in-context learners Kumo Tabular S / M / L and TabICLv2 (row encoder output, every in-context layer, final norm).

Per layer we record (a) supervised probes (linear, kNN) with the layer chosen on validation and scored with the
official RelBench evaluator, and (b) label-free geometry on a fixed sample of test rows: effective rank (RankMe),
TwoNN intrinsic dimension, anisotropy (mean pairwise cosine), k-means silhouette, plus linear CKA between all layers of
a model and between the models' layers on the same rows. The question the geometry answers: can the best layer be
picked without labels, and do different models build the same representation?
"""
from __future__ import annotations

import time
from typing import Any

import numpy as np
import pandas as pd

from fer.relbench_layers import Subgrapher, classification_module, probe_scores_multi, safe_auc, time_split
from fer.relbench_tabular_layers import BlockTap, row_features


# ----------------------------------------------------------------------------------------------- geometry
def effective_rank(E: np.ndarray) -> float:
    s = np.linalg.svd(E - E.mean(0), compute_uv=False)
    p = s / max(s.sum(), 1e-12)
    p = p[p > 0]
    return float(np.exp(-(p * np.log(p)).sum()))


def twonn_id(E: np.ndarray) -> float:
    from sklearn.neighbors import NearestNeighbors

    d, _ = NearestNeighbors(n_neighbors=3).fit(E).kneighbors(E)
    r1, r2 = d[:, 1], d[:, 2]
    ok = r1 > 1e-12
    mu = r2[ok] / r1[ok]
    mu = mu[np.isfinite(mu) & (mu > 1)]
    return float(len(mu) / np.log(mu).sum()) if len(mu) else float("nan")


def anisotropy(E: np.ndarray, n: int = 1000, seed: int = 0) -> float:
    rng = np.random.default_rng(seed)
    X = E[rng.choice(len(E), min(n, len(E)), replace=False)]
    X = X / (np.linalg.norm(X, axis=1, keepdims=True) + 1e-12)
    C = X @ X.T
    return float((C.sum() - np.trace(C)) / (len(X) * (len(X) - 1)))


def kmeans_silhouette(E: np.ndarray, k: int = 8, seed: int = 0) -> float:
    from sklearn.cluster import KMeans
    from sklearn.metrics import silhouette_score

    Z = (E - E.mean(0)) / (E.std(0) + 1e-9)
    lab = KMeans(k, n_init=4, random_state=seed).fit_predict(Z)
    return float(silhouette_score(Z, lab)) if len(np.unique(lab)) > 1 else float("nan")


def linear_cka(A: np.ndarray, B: np.ndarray) -> float:
    A = A - A.mean(0); B = B - B.mean(0)
    hsic = np.linalg.norm(A.T @ B) ** 2
    return float(hsic / (np.linalg.norm(A.T @ A) * np.linalg.norm(B.T @ B) + 1e-12))


def geometry(E: np.ndarray) -> dict[str, float]:
    E = E.astype(np.float64)
    return {"effective_rank": effective_rank(E), "intrinsic_dim": twonn_id(E), "anisotropy": anisotropy(E),
            "kmeans_silhouette": kmeans_silhouette(E)}


# ----------------------------------------------------------------------------------------------- taps
class SdmTap:
    """Hooks on an sdm in-context model (Kumo Tabular, TabICLv2): row encoder, each ICL layer, final norm."""

    def __init__(self, inner):
        self.calls: dict[str, list] = {}
        self.handles = [inner.row_embedding.register_forward_hook(self._hook("row_emb"))]
        for i, layer in enumerate(inner.icl_block.layers):
            self.handles.append(layer.register_forward_hook(self._hook(f"icl_{i:02d}")))
        self.handles.append(inner.icl_block.norm.register_forward_hook(self._hook("final")))

    def _hook(self, name):
        def hook(mod, args, out):
            o = out[0] if isinstance(out, tuple) else out
            self.calls.setdefault(name, []).append(o.detach().float().cpu())
        return hook

    def take(self, n: int) -> tuple[dict[str, np.ndarray], dict[str, int]]:
        res, ncalls = {}, {}
        for name, lst in self.calls.items():
            ncalls[name] = len(lst)
            rows = [x.reshape(-1, x.shape[-1]) for x in lst]  # sub-chunks of the queries, in order
            res[name] = np.concatenate([r.numpy() for r in rows], 0)[-n:]
        self.calls = {}
        return res, ncalls

    def remove(self):
        for h in self.handles:
            h.remove()


def model_layer_embeddings(spec: str, X_ctx: pd.DataFrame, y_ctx: np.ndarray, X_sets: list[pd.DataFrame],
                           chunk: int = 1000) -> tuple[list[dict[str, np.ndarray]], list[np.ndarray | None], dict]:
    """Per-layer embeddings of every row of each frame in ``X_sets`` and, for a binary 0/1 context, P(y=1)."""
    from tabfm_auto.models import get_model
    from tabfm_auto.models.sdm_wrapper import SdmEstimator

    n_cls = len(np.unique(y_ctx))
    task_type = "binary" if n_cls <= 2 else "multiclass"
    name, _, opts = spec.partition(":")
    kw = dict(o.split("=", 1) for o in opts.split(",") if "=" in o)
    if name.startswith("kumo-tabular") or name.startswith("tabiclv2"):
        # built directly: sdm downloads its own checkpoint; the registry's status check does not know about it
        fam = "kumo-tabular" if name.startswith("kumo-tabular") else "tabiclv2"
        size = {"s": "small", "m": "medium", "l": "large"}.get(name.rsplit("-", 1)[-1]) if fam == "kumo-tabular" else None
        est = SdmEstimator(fam, task_type, size=size, n_estimators=int(kw.get("n_estimators", 1)), device=kw.get("device", "cuda"))
    else:
        est = get_model(spec, task_type)
    est.fit(X_ctx, pd.Series(y_ctx) if not spec.startswith("tabpfn") else y_ctx)
    is_tabpfn = spec.startswith("tabpfn")
    tap = BlockTap(est.models_[0]) if is_tabpfn else SdmTap(classification_module(est.model))
    binary = n_cls == 2 and set(np.unique(y_ctx)) <= {0, 1}
    outs, scores, diag = [], [], {"calls_per_chunk": {}}
    try:
        for X in X_sets:
            parts: dict[str, list] = {}
            sc = []
            for s in range(0, len(X), chunk):
                xc = X.iloc[s:s + chunk]
                proba = est.predict_proba(xc)
                if is_tabpfn:
                    got = tap.take(len(xc))
                else:
                    got, nc = tap.take(len(xc))
                    diag["calls_per_chunk"] = nc
                for k, v in got.items():
                    parts.setdefault(k, []).append(v)
                if binary:
                    sc.append(np.asarray(proba)[:, list(getattr(est, "classes_", [0, 1])).index(1)])
            if not parts:
                raise RuntimeError(f"{spec}: no layer was recorded")
            outs.append({k: np.concatenate(v, 0) for k, v in parts.items()})
            scores.append(np.concatenate(sc) if sc else None)
    finally:
        tap.remove()
    return outs, scores, diag


def _order(layers):
    head = [L for L in ("raw_features", "row_emb") if L in layers]
    mid = sorted(L for L in layers if L.startswith(("block_", "icl_")))
    return head + mid + (["final"] if "final" in layers else [])


def run_task(dataset: str, task_name: str, specs: dict[str, str], targets=("random", "kmeans", "label"),
             n_ctx: int = 3000, n_train: int = 4000, k_children: int = 50, n_geom: int = 2000, seed: int = 0,
             max_eval: int | None = 5000, log: Any = None, ctx_split: str = "time") -> dict[str, Any]:
    import relbench
    from scipy.stats import spearmanr
    from sklearn.cluster import KMeans
    from sklearn.metrics import roc_auc_score
    from sklearn.preprocessing import StandardScaler

    task = relbench.load_dataset(dataset).load_task(task_name)
    db = task.get_db(upto_test_timestamp=False)
    ent_col, time_col, tgt = task.entity_col, task.time_col, task.target_col
    tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
    va = task.get_table("val", mask_input_cols=False).df.reset_index(drop=True)
    te = task.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
    rng = np.random.default_rng(seed)
    if max_eval and len(va) > max_eval:
        va = va.iloc[np.sort(rng.choice(len(va), max_eval, replace=False))].reset_index(drop=True)
    sub = Subgrapher(db, task.entity_table, k_children=k_children)
    if ctx_split == "time":  # probe-train rows strictly later than every context row, like val/test (see time_split)
        early, late, split_info = time_split(tr, time_col, n_train)
        nc = min(n_ctx, len(early))
        ctx = early.iloc[rng.permutation(len(early))[:nc]].reset_index(drop=True)
        prb = late.iloc[rng.permutation(len(late))[:n_train]].reset_index(drop=True)
    else:
        perm = rng.permutation(len(tr))
        nc = min(n_ctx, len(tr) // 2)
        ctx, prb = tr.iloc[perm[:nc]].reset_index(drop=True), tr.iloc[perm[nc:nc + n_train]].reset_index(drop=True)
        split_info = {"ctx_split": "random"}
    F = {k: row_features(sub, d[ent_col].to_numpy(), d[time_col].to_numpy(), k_children)
         for k, d in (("ctx", ctx), ("prb", prb), ("va", va), ("te", te))}
    geo_idx = np.sort(rng.choice(len(te), min(n_geom, len(te)), replace=False))
    ytr = prb[tgt].to_numpy().astype(int)
    res: dict[str, Any] = {"dataset": dataset, "task": task_name, "n_features": int(F["ctx"].shape[1]), "n_ctx": nc,
                           "n_probe_train": len(prb), "n_val": len(va), "n_test": len(te), **split_info, "models": {}}
    geo_store: dict[tuple[str, str], dict[str, np.ndarray]] = {}
    for mname, spec in specs.items():
        res["models"][mname] = {}
        for target in targets:
            if target == "random":
                y_ctx = rng.integers(0, 2, nc)
            elif target == "kmeans":
                y_ctx = KMeans(8, n_init=4, random_state=seed).fit_predict(StandardScaler().fit_transform(F["ctx"]))
            else:
                y_ctx = ctx[tgt].to_numpy().astype(int)
            m: dict[str, Any] = {"layers": {}}
            t0 = time.time()
            try:
                (Etr, Eva, Ete), (s_tr, s_va, s_te), diag = model_layer_embeddings(spec, F["ctx"], y_ctx, [F["prb"], F["va"], F["te"]])
            except Exception as e:
                m["error"] = f"{type(e).__name__}: {str(e)[:400]}"
                res["models"][mname][target] = m
                if log is not None:
                    log.event("model_layers_error", model=mname, target=target, error=m["error"])
                continue
            m["seconds"] = round(time.time() - t0, 1)
            m["diag"] = diag
            if target == "label" and s_te is not None:
                m["icl_head"] = {"val_auroc": float(roc_auc_score(va[tgt], s_va)),
                                 "test": {k: float(v) for k, v in task.evaluate(s_te).items()},
                                 "probe_train_auroc": safe_auc(ytr, s_tr)}
            Etr["raw_features"], Eva["raw_features"], Ete["raw_features"] = F["prb"].to_numpy(), F["va"].to_numpy(), F["te"].to_numpy()
            for L in _order(Etr):
                pv, pt = probe_scores_multi(Etr[L], ytr, [Eva[L], Ete[L]])
                G = Ete[L][geo_idx]
                if L == "raw_features":  # mixed units (counts, days, means): compare shapes, not scales
                    G = (G - G.mean(0)) / (G.std(0) + 1e-9)
                g = geometry(G)
                m["layers"][L] = {"geometry": g, **{p: {"val_auroc": float(roc_auc_score(va[tgt], pv[p])),
                                                        "test_auroc": float(task.evaluate(pt[p]).get("roc_auc", np.nan))} for p in pv}}
                geo_store.setdefault((mname, target), {})[L] = Ete[L][geo_idx].astype(np.float32)
                if log is not None:
                    log.event("layer", model=mname, target=target, layer=L, linear_test=m["layers"][L]["linear"]["test_auroc"],
                              knn_test=m["layers"][L]["knn"]["test_auroc"], **g)
            model_layers = [L for L in m["layers"] if L != "raw_features"]
            for p in ("linear", "knn"):
                best = max(model_layers, key=lambda L: m["layers"][L][p]["val_auroc"])
                m[f"best_by_val_{p}"] = {"layer": best, **m["layers"][best][p]}
            # can a label-free geometric statistic pick the layer?
            y_layer = np.array([m["layers"][L]["linear"]["test_auroc"] for L in model_layers])
            m["geometry_vs_probe"] = {}
            for stat in ("effective_rank", "intrinsic_dim", "anisotropy", "kmeans_silhouette"):
                x = np.array([m["layers"][L]["geometry"][stat] for L in model_layers])
                rho = spearmanr(x, y_layer).statistic if np.isfinite(x).all() and len(set(x)) > 1 else float("nan")
                pick_hi = model_layers[int(np.nanargmax(x))]; pick_lo = model_layers[int(np.nanargmin(x))]
                m["geometry_vs_probe"][stat] = {"spearman_with_linear_test": float(rho),
                                                "argmax_layer": pick_hi, "argmax_test": m["layers"][pick_hi]["linear"]["test_auroc"],
                                                "argmin_layer": pick_lo, "argmin_test": m["layers"][pick_lo]["linear"]["test_auroc"]}
            # within-model CKA between layers
            store = geo_store[(mname, target)]
            names = [L for L in _order(store) if L != "raw_features"]
            m["cka_layers"] = {"layers": names, "matrix": [[linear_cka(store[a], store[b]) for b in names] for a in names]}
            res["models"][mname][target] = m
    # across models: CKA of each model's val-selected layer and final layer, per target
    res["cka_models"] = {}
    for target in targets:
        reps = {}
        for mname in specs:
            m = res["models"][mname].get(target, {})
            if "layers" not in m or (mname, target) not in geo_store:
                continue
            st = geo_store[(mname, target)]
            reps[f"{mname}:best"] = st[m["best_by_val_linear"]["layer"]]
            last = _order(st)[-1]
            reps[f"{mname}:last"] = st[last]
        reps["raw_features"] = F["te"].to_numpy()[geo_idx]
        keys = list(reps)
        res["cka_models"][target] = {"keys": keys, "matrix": [[linear_cka(reps[a], reps[b]) for b in keys] for a in keys]}
    return res
