"""Stage D: how well does each signal separate correct from wrong SQL candidates, and what does it buy as a selector?

  AUROC (pooled, and mean within questions that have both correct and wrong candidates)
  best-of-N accuracy: pick each question's top-scored candidate (ties: self-consistency, then greedy)
  ECE of the raw score and after isotonic calibration fitted on the other half of the questions (2-fold by question)
  stacked logistic regression over cheap signals (+ GLiClass taxonomy), with and without the judge
  cascade: stacked cheap score; send the middle band to the judge; accuracy vs share of judge calls
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.isotonic import IsotonicRegression
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import roc_auc_score
from sklearn.preprocessing import StandardScaler


def auroc(y, s) -> float:
    y, s = np.asarray(y), np.asarray(s, float)
    ok = np.isfinite(s)
    return float(roc_auc_score(y[ok], s[ok])) if ok.sum() > 1 and len(set(y[ok])) == 2 else float("nan")


def within_auroc(df: pd.DataFrame, col: str) -> float:
    vals = [auroc(g["correct"], g[col]) for _, g in df.groupby("qid") if g["correct"].nunique() == 2]
    return float(np.nanmean(vals)) if vals else float("nan")


def best_of_n(df: pd.DataFrame, col: str) -> float:
    picks = []
    for _, g in df.groupby("qid"):
        g = g.assign(_tie1=g.get("sig_self_consistency", 0), _tie2=g["greedy"].astype(int))
        picks.append(bool(g.sort_values([col, "_tie1", "_tie2"], ascending=False).iloc[0]["correct"]))
    return float(np.mean(picks))


def ece(y, p, bins: int = 10) -> float:
    y, p = np.asarray(y, float), np.clip(np.asarray(p, float), 0, 1)
    edges = np.linspace(0, 1, bins + 1)
    idx = np.clip(np.digitize(p, edges) - 1, 0, bins - 1)
    return float(sum(abs(p[idx == b].mean() - y[idx == b].mean()) * (idx == b).mean() for b in range(bins) if (idx == b).any()))


def folds(df: pd.DataFrame, k: int = 2, seed: int = 0) -> np.ndarray:
    q = np.array(sorted(df["qid"].unique()))
    rng = np.random.default_rng(seed)
    f = dict(zip(q, rng.permutation(len(q)) % k))
    return df["qid"].map(f).to_numpy()


def cross_fit(df: pd.DataFrame, cols: list[str], k: int = 2) -> np.ndarray:
    """Out-of-fold stacked P(correct) from a logistic regression over `cols` (fit on the other questions)."""
    X = df[cols].to_numpy(float)
    X = np.where(np.isfinite(X), X, np.nanmin(np.where(np.isfinite(X), X, np.nan), 0))
    y = df["correct"].to_numpy(int)
    fo = folds(df, k)
    out = np.zeros(len(df))
    for i in range(k):
        tr, te = fo != i, fo == i
        sc = StandardScaler().fit(X[tr])
        m = LogisticRegression(C=1.0, max_iter=2000).fit(sc.transform(X[tr]), y[tr])
        out[te] = m.predict_proba(sc.transform(X[te]))[:, 1]
    return out


def calibrated(df: pd.DataFrame, col: str, k: int = 2) -> np.ndarray:
    fo = folds(df, k)
    out = np.zeros(len(df))
    s, y = df[col].to_numpy(float), df["correct"].to_numpy(int)
    for i in range(k):
        tr, te = fo != i, fo == i
        out[te] = IsotonicRegression(out_of_bounds="clip").fit(s[tr], y[tr]).predict(s[te])
    return out


def cascade(df: pd.DataFrame, cheap: str, judge: str, shares=(0.0, 0.1, 0.2, 0.3, 0.5, 1.0)) -> list[dict]:
    """Send the `share` of candidates whose cheap score is closest to its median-uncertainty point (0.5 after
    calibration) to the judge; the rest keep the cheap score. Report AUROC and best-of-N for each share."""
    p = df[cheap].to_numpy(float)
    dist = np.abs(p - 0.5)
    order = np.argsort(dist)
    rows = []
    for sh in shares:
        n = int(round(sh * len(df)))
        mix = p.copy()
        idx = order[:n]
        mix[idx] = df[judge].to_numpy(float)[idx]
        d = df.assign(_mix=mix)
        rows.append({"judge_share": sh, "auroc": auroc(d["correct"], d["_mix"]), "best_of_n": best_of_n(d, "_mix")})
    return rows


def summarize(df: pd.DataFrame) -> dict:
    df = df[df["gold_ok"]].copy()
    sigs = [c for c in df.columns if c.startswith("sig_")]
    tax = [c for c in df.columns if c.startswith("tax_")]
    cheap = [c for c in ["sig_exec_ok", "sig_nonempty", "sig_logprob", "sig_self_consistency"] if c in df]
    if "sig_gliclass" in df:
        df["stack_cheap"] = cross_fit(df, cheap)
        df["stack_cheap_gli"] = cross_fit(df, cheap + ["sig_gliclass"] + tax)
    if "sig_nli" in df:
        df["stack_cheap_gli_nli"] = cross_fit(df, cheap + [c for c in ["sig_gliclass", "sig_nli"] if c in df] + tax)
    if "sig_judge" in df:
        df["stack_all"] = cross_fit(df, cheap + [c for c in ["sig_gliclass", "sig_nli", "sig_judge"] if c in df] + tax)
    stacks = [c for c in df.columns if c.startswith("stack_")]
    res = {"n_questions": int(df["qid"].nunique()), "n_candidates": int(len(df)),
           "candidates_correct_share": float(df["correct"].mean()),
           "greedy_accuracy": float(df[df["greedy"]]["correct"].mean()),
           "oracle_accuracy": float(df.groupby("qid")["correct"].any().mean()),
           "signals": {}}
    for c in sigs + stacks:
        r = {"auroc": auroc(df["correct"], df[c]), "within_question_auroc": within_auroc(df, c),
             "best_of_n": best_of_n(df, c)}
        s = df[c].to_numpy(float)
        if np.nanmin(s) >= 0 and np.nanmax(s) <= 1:
            r["ece_raw"] = ece(df["correct"], s)
        r["ece_isotonic"] = ece(df["correct"], calibrated(df, c))
        res["signals"][c] = r
    if "sig_judge" in df and "stack_cheap_gli" in df:
        res["cascade"] = cascade(df, "stack_cheap_gli", "sig_judge")
    res["_frame"] = df
    return res
