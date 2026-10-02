"""Bradley-Terry Elo as used by TabArena / TabFM-Auto (Bradley & Terry 1952; Hunter 2004 MM algorithm).

Input: a long dataframe with columns ``method, task, error`` (one row per method x task, a task being a
dataset-split; lower error is better). Each pair of methods plays one game per shared task; the lower error wins,
ties count half. Ratings are shifted so that ``anchor`` has ``anchor_rating`` (TabArena: RandomForest (default) = 1000).
"""
from __future__ import annotations

import numpy as np
import pandas as pd


def pairwise_wins(df: pd.DataFrame) -> tuple[list[str], np.ndarray]:
    piv = df.pivot_table(index="task", columns="method", values="error")
    methods = list(piv.columns)
    E = piv.to_numpy()
    n = len(methods)
    W = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            both = ~np.isnan(E[:, i]) & ~np.isnan(E[:, j])
            W[i, j] = (E[both, i] < E[both, j]).sum() + 0.5 * (E[both, i] == E[both, j]).sum()
    return methods, W


def bradley_terry(W: np.ndarray, iters: int = 5000, tol: float = 1e-10) -> np.ndarray:
    """Hunter's MM algorithm. Returns log-strengths with geometric mean 1 (log-mean 0)."""
    n = W.shape[0]
    p = np.ones(n)
    N = W + W.T
    w = W.sum(1)
    for _ in range(iters):
        denom = (N / (p[:, None] + p[None, :] + 1e-300)).sum(1)
        p_new = np.where(w > 0, w / np.clip(denom, 1e-300, None), p)
        p_new = p_new / np.exp(np.mean(np.log(p_new + 1e-300)))
        done = np.max(np.abs(np.log(p_new + 1e-300) - np.log(p + 1e-300))) < tol
        p = p_new
        if done:
            break
    return np.log(p + 1e-300)


def elo_ratings(df: pd.DataFrame, anchor: str | None = None, anchor_rating: float = 1000.0, scale: float = 400.0,
                n_bootstrap: int = 0, seed: int = 0) -> pd.DataFrame:
    """Elo per method, optionally with a bootstrap-over-tasks 95% interval (TabArena uses 100 rounds)."""
    methods, W = pairwise_wins(df)
    elo = scale / np.log(10) * bradley_terry(W)
    if anchor is not None and anchor in methods:
        elo += anchor_rating - elo[methods.index(anchor)]
    out = pd.DataFrame({"method": methods, "elo": elo})
    if n_bootstrap:
        rng = np.random.default_rng(seed)
        tasks = df.task.unique()
        boots = []
        for _ in range(n_bootstrap):
            sample = rng.choice(tasks, len(tasks), replace=True)
            dfb = pd.concat([df[df.task == t].assign(task=f"{t}#{k}") for k, t in enumerate(sample)])
            m2, W2 = pairwise_wins(dfb)
            e2 = scale / np.log(10) * bradley_terry(W2)
            if anchor is not None and anchor in m2:
                e2 += anchor_rating - e2[m2.index(anchor)]
            boots.append(pd.Series(e2, index=m2))
        B = pd.concat(boots, axis=1)
        out["elo_lo"] = B.quantile(0.025, axis=1).reindex(methods).to_numpy()
        out["elo_hi"] = B.quantile(0.975, axis=1).reindex(methods).to_numpy()
    return out.sort_values("elo", ascending=False).reset_index(drop=True)
