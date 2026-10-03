"""The VFD harness, SQL path (claim C1). One question -> one SQL, with each component switchable for ablations.

  G1 link     GLiClass / GLiNER / reranker ranks the columns; the LLM sees only the top-k columns (plus keys)
  L1 write    the LLM writes n candidate SQL (one greedy + n-1 sampled)
  G2 triage   GLiClass scores each candidate against an error taxonomy; risky ones go back to the LLM once with the flag
  X  execute  every candidate runs read-only
  V  verify   pick by a verifier score: self-consistency, a calibrated stack of cheap signals, or a cascade that sends
              the uncertain band to an LLM judge
Configs (ablations): A = L1 greedy only; B = A + G1; C = B + G2; D = C + V (n candidates, cheap verifier);
F = D + cascade judge. Everything is logged per question: tokens, LLM calls, seconds, chosen SQL, correctness.
"""
from __future__ import annotations

import time
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

import numpy as np

from . import bird
from .llm import Chat, extract_sql
from .schema_link import column_label

RISK_LABELS = ["uses a wrong column", "misses a filter condition", "joins the wrong tables", "wrong aggregation",
               "returns extra columns"]


@dataclass
class Config:
    name: str
    link: str | None = None          # None | "gliclass" | "gliner" | "rerank"
    link_k: int = 20
    triage: bool = False
    triage_threshold: float = 0.5
    n: int = 1
    verifier: str = "greedy"          # "greedy" | "self_consistency" | "stack" | "cascade"
    cascade_band: tuple[float, float] = (0.3, 0.7)


CONFIGS = {
    "A": Config("A"),
    "B": Config("B", link="gliclass"),
    "C": Config("C", link="gliclass", triage=True),
    "D": Config("D", link="gliclass", triage=True, n=8, verifier="stack"),
    "F": Config("F", link="gliclass", triage=True, n=8, verifier="cascade"),
    "SC": Config("SC", n=8, verifier="self_consistency"),   # the standard test-time-scaling baseline, no small models
}


@dataclass
class Tools:
    chat: Chat
    judge: Chat | None = None
    gliclass: Any = None             # vfd.signals.GLiClassScorer
    linker: Any = None               # object with .scores(query, labels)
    stack: Any = None                # fitted (scaler, logistic regression, feature names) from the P0 study
    usage: Counter = field(default_factory=Counter)


def link_schema(schema: bird.Schema, question: str, tools: Tools, k: int) -> set[tuple[str, str]]:
    cols = schema.columns
    s = tools.linker.scores(question, [column_label(c) for c in cols])
    keep = {(cols[j].table.lower(), cols[j].name.lower()) for j in np.argsort(-s)[:k]}
    # keep join keys of every kept table so the LLM can still join
    tabs = {t for t, _ in keep}
    for a, b, c, d in schema.fks:
        if a.lower() in tabs or c.lower() in tabs:
            keep.add((a.lower(), b.lower()))
            if d:
                keep.add((c.lower(), d.lower()))
    for t, tcols in schema.tables.items():
        if t.lower() in tabs:
            for c in tcols:
                if c.name.lower() in ("id", f"{t.lower()}_id") or c.name.lower().endswith("id"):
                    keep.add((t.lower(), c.name.lower()))
    return keep


def _messages(q: bird.Question, schema_text: str, feedback: str = "") -> list[dict]:
    from .candidates import SYSTEM

    user = f"Schema:\n{schema_text}\n\nEvidence: {q.evidence or '(none)'}\nQuestion: {q.question}"
    if feedback:
        user += f"\n\nA reviewer flagged a previous attempt:\n{feedback}\nWrite a corrected query."
    return [{"role": "system", "content": SYSTEM}, {"role": "user", "content": user}]


def answer(q: bird.Question, db_path, schema: bird.Schema, cfg: Config, tools: Tools) -> dict:
    t0 = time.time()
    log: dict[str, Any] = {"qid": q.qid, "config": cfg.name, "llm_calls": 0, "judge_calls": 0, "triage_revisions": 0}
    cols = link_schema(schema, f"{q.question} {q.evidence}", tools, cfg.link_k) if cfg.link else None
    schema_text = bird.render_schema(schema, cols)
    log["schema_chars"] = len(schema_text)
    msgs = _messages(q, schema_text)
    samples = tools.chat.samples(msgs, n=1, temperature=0.0)
    log["llm_calls"] += 1
    if cfg.n > 1:
        samples += tools.chat.samples(msgs, n=cfg.n - 1, temperature=0.8)
        log["llm_calls"] += 1
    cands = []
    for k, smp in enumerate(samples):
        sql = extract_sql(smp.text)
        r = bird.execute(db_path, sql)
        cands.append({"k": k, "sql": sql, "res": r, "lp": smp.mean_logprob, "tokens": smp.n_tokens})
    log["completion_tokens"] = sum(c["tokens"] for c in cands)
    if cfg.triage and tools.gliclass is not None:
        texts = [f"Question: {q.question}\nEvidence: {q.evidence}\nSQL: {c['sql']}\nResult: {bird.preview(c['res'])}"
                 for c in cands]
        risk = tools.gliclass.scores(texts, RISK_LABELS)
        for c, rv in zip(cands, risk):
            c["risk"] = float(rv.max())
            if rv.max() > cfg.triage_threshold or not c["res"].ok:
                flag = (c["res"].error if not c["res"].ok else RISK_LABELS[int(rv.argmax())])
                fb = f"SQL: {c['sql']}\nIssue: {flag}"
                smp = tools.chat.samples(_messages(q, schema_text, fb), n=1, temperature=0.0)[0]
                log["llm_calls"] += 1
                log["triage_revisions"] += 1
                sql = extract_sql(smp.text)
                c.update(sql=sql, res=bird.execute(db_path, sql), lp=smp.mean_logprob, revised=True)
    # self-consistency over result sets
    keys = Counter(c["res"].key for c in cands if c["res"].ok)
    for c in cands:
        c["sc"] = keys.get(c["res"].key, 0) / len(cands) if c["res"].ok else 0.0
    if cfg.verifier == "greedy" or len(cands) == 1:
        pick = cands[0]
    elif cfg.verifier == "self_consistency":
        pick = max(cands, key=lambda c: (c["sc"], c["k"] == 0))
    else:
        feats = _stack_features(q, cands, tools)
        scaler, lr, _ = tools.stack
        p = lr.predict_proba(scaler.transform(feats))[:, 1]
        for c, pi in zip(cands, p):
            c["p"] = float(pi)
        if cfg.verifier == "cascade" and tools.judge is not None:
            lo, hi = cfg.cascade_band
            for c in cands:
                if lo <= c["p"] <= hi:
                    c["p"] = _judge(q, schema_text, c, tools)
                    log["judge_calls"] += 1
        pick = max(cands, key=lambda c: (c["p"], c["sc"], c["k"] == 0))
    log.update(sql=pick["sql"], exec_ok=pick["res"].ok, result_key=pick["res"].key, seconds=round(time.time() - t0, 2))
    return log


def _stack_features(q, cands, tools) -> np.ndarray:
    from .signals import HYPOTHESES

    texts = [f"Question: {q.question}\nEvidence: {q.evidence or '(none)'}\nSQL: {c['sql']}\nResult: {bird.preview(c['res'])}"
             for c in cands]
    gli = np.stack([tools.gliclass.scores(texts, [h])[:, 0] for h in HYPOTHESES], 1).mean(1) if tools.gliclass else 0
    lp = np.array([c["lp"] if np.isfinite(c["lp"]) else -5.0 for c in cands])
    ok = np.array([float(c["res"].ok) for c in cands])
    ne = np.array([float(c["res"].ok and len(c["res"].rows) > 0) for c in cands])
    sc = np.array([c["sc"] for c in cands])
    return np.column_stack([ok, ne, lp, sc, gli])


def _judge(q, schema_text, c, tools) -> float:
    from .signals import JUDGE_SYSTEM

    msgs = [{"role": "system", "content": JUDGE_SYSTEM},
            {"role": "user", "content": f"Schema:\n{schema_text}\n\nQuestion: {q.question}\nEvidence: {q.evidence}\n"
                                        f"SQL: {c['sql']}\nResult: {bird.preview(c['res'])}\n\nIs the query correct? Answer Yes or No."}]
    return tools.judge.yes_prob(msgs)[0]


def fit_stack(signals_csv: str):
    """Fit the cheap verifier stack (exec_ok, nonempty, logprob, self-consistency, GLiClass) on the P0 candidates."""
    import pandas as pd
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler

    df = pd.read_csv(signals_csv)
    df = df[df["gold_ok"]]
    names = ["sig_exec_ok", "sig_nonempty", "sig_logprob", "sig_self_consistency", "sig_gliclass"]
    X = df[names].to_numpy(float)
    X = np.where(np.isfinite(X), X, -5.0)
    sc = StandardScaler().fit(X)
    lr = LogisticRegression(max_iter=2000).fit(sc.transform(X), df["correct"].astype(int))
    return sc, lr, names
