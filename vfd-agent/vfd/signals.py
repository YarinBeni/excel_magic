"""Stage B/C of the verifier study: cheap and expensive correctness signals for each SQL candidate.

  cheap, no model:   exec_ok, nonempty, mean token logprob, self-consistency (share of the question's candidates with the
                     same result set)
  small encoders:    GLiClass (hypothesis averaged over paraphrases + an error-taxonomy label vector), an NLI
                     cross-encoder (deberta-v3-large-zeroshot-v2.0)
  LLM judge:         a second LLM family, P(yes) from first-token logprobs, sees the schema too
Every scorer records seconds per item so latency can be compared.
"""
from __future__ import annotations

import time
from collections import Counter

import numpy as np
import pandas as pd

HYPOTHESES = ["The SQL query correctly answers the question.",
              "The query result is the correct answer to the question.",
              "This SQL query does what the question asks."]
ERROR_LABELS = ["uses a wrong column", "misses a filter condition", "joins the wrong tables", "wrong aggregation",
                "returns extra columns", "wrong ordering or limit", "the result is empty", "the query is correct"]


def candidate_text(row: dict, q: dict) -> str:
    return (f"Question: {q['question']}\nEvidence: {q['evidence'] or '(none)'}\nSQL: {row['sql']}\n"
            f"Result: {row['preview']}")


def cheap_signals(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["sig_exec_ok"] = df["exec_ok"].astype(float)
    df["sig_nonempty"] = ((df["exec_ok"]) & (df["n_rows"] > 0)).astype(float)
    df["sig_logprob"] = df["mean_logprob"].fillna(df["mean_logprob"].min())
    sc = []
    for _, g in df.groupby("qid", sort=False):
        keys = Counter(k for k in g["result_key"] if k != "ERR")
        n = len(g)
        sc.extend((keys.get(k, 0) / n) if k != "ERR" else 0.0 for k in g["result_key"])
    df["sig_self_consistency"] = sc
    return df


class GLiClassScorer:
    def __init__(self, model_id: str = "knowledgator/gliclass-large-v3.0", device: str = "cuda:0", batch_size: int = 16):
        from gliclass import GLiClassModel, ZeroShotClassificationPipeline
        from transformers import AutoTokenizer

        model = GLiClassModel.from_pretrained(model_id)
        tok = AutoTokenizer.from_pretrained(model_id, add_prefix_space=True)
        self.pipe = ZeroShotClassificationPipeline(model, tok, classification_type="multi-label", device=device)
        self.batch_size = batch_size

    def scores(self, texts: list[str], labels: list[str]) -> np.ndarray:
        """[n_texts, n_labels] label scores; threshold 0 so every label is returned."""
        out = np.zeros((len(texts), len(labels)), np.float32)
        idx = {lab: j for j, lab in enumerate(labels)}
        for s in range(0, len(texts), self.batch_size):
            res = self.pipe(texts[s:s + self.batch_size], labels, threshold=0.0)
            if res and isinstance(res[0], dict):  # single-text form
                res = [res]
            for i, per in enumerate(res):
                for d in per:
                    if d["label"] in idx:
                        out[s + i, idx[d["label"]]] = d["score"]
        return out


class NLIScorer:
    def __init__(self, model_id: str = "MoritzLaurer/deberta-v3-large-zeroshot-v2.0", device: int = 0, batch_size: int = 16):
        from transformers import pipeline

        self.pipe = pipeline("zero-shot-classification", model=model_id, device=device)
        self.batch_size = batch_size

    def scores(self, texts: list[str], hypotheses: list[str]) -> np.ndarray:
        out = np.zeros((len(texts), len(hypotheses)), np.float32)
        res = self.pipe(texts, candidate_labels=hypotheses, hypothesis_template="{}", multi_label=True,
                        batch_size=self.batch_size)
        if isinstance(res, dict):
            res = [res]
        for i, r in enumerate(res):
            m = dict(zip(r["labels"], r["scores"]))
            out[i] = [m[h] for h in hypotheses]
        return out


def encoder_signals(df: pd.DataFrame, questions: dict[int, dict], gliclass: GLiClassScorer | None,
                    nli: NLIScorer | None) -> tuple[pd.DataFrame, dict]:
    df = df.copy()
    texts = [candidate_text(r, questions[r["qid"]]) for r in df.to_dict("records")]
    lat = {}
    if gliclass is not None:
        t0 = time.time()
        per_h = np.stack([gliclass.scores(texts, [h])[:, 0] for h in HYPOTHESES], 1)  # one hypothesis per call
        df["sig_gliclass"] = per_h.mean(1)
        lat["gliclass_s_per_item"] = (time.time() - t0) / max(1, len(texts)) / len(HYPOTHESES)
        t0 = time.time()
        tax = gliclass.scores(texts, ERROR_LABELS)
        lat["gliclass_taxonomy_s_per_item"] = (time.time() - t0) / max(1, len(texts))
        for j, lab in enumerate(ERROR_LABELS):
            df[f"tax_{lab.replace(' ', '_')}"] = tax[:, j]
    if nli is not None:
        t0 = time.time()
        df["sig_nli"] = nli.scores(texts, HYPOTHESES).mean(1)
        lat["nli_s_per_item"] = (time.time() - t0) / max(1, len(texts))
    return df, lat


JUDGE_SYSTEM = ("You review SQL written by another analyst. Decide if the SQLite query correctly answers the question "
                "on this database, using the evidence. Answer with one word: Yes or No.")


def judge_signal(df: pd.DataFrame, questions: dict[int, dict], schemas: dict[str, str], chat, workers: int = 32
                 ) -> tuple[pd.DataFrame, dict]:
    from .llm import pmap

    df = df.copy()
    rows = df.to_dict("records")

    def one(r):
        q = questions[r["qid"]]
        msgs = [{"role": "system", "content": JUDGE_SYSTEM},
                {"role": "user", "content": f"Schema:\n{schemas[r['db_id']]}\n\n{candidate_text(r, q)}\n\nIs the query correct? "
                                            "Answer Yes or No."}]
        return chat.yes_prob(msgs)

    t0 = time.time()
    res = pmap(one, rows, workers=workers)
    df["sig_judge"] = [p for p, _ in res]
    fallback = float(np.mean([t.startswith("fallback:") for _, t in res])) if res else 0.0
    return df, {"judge_s_per_item_wall": (time.time() - t0) / max(1, len(rows)), "judge_workers": workers,
                "judge_fallback_share": fallback}
