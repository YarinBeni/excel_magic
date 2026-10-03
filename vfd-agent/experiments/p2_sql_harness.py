"""P2 / claim C1: the harness on BIRD, ablation A, B, C, D, F and the self-consistency baseline SC.

Questions: Arcwise-Plat (498, fully corrected). The verifier stack used by D/F is fitted on the P0 candidates of the
OTHER half of the questions (2-fold by question), so no question is scored by a stack that saw it.

  python experiments/p2_sql_harness.py --run DIR --signals P0RUN/signals.csv --configs A,B,C,D,F,SC
needs OPENAI_BASE_URL (generator) and, for F, JUDGE_BASE_URL (judge LLM).
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402
from vfd.harness import CONFIGS, Tools, answer  # noqa: E402
from vfd.llm import Chat, pmap  # noqa: E402


def fit_half_stacks(signals_csv: str, qids_by_fold: dict[int, int]):
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler

    df = pd.read_csv(signals_csv)
    df = df[df["gold_ok"]]
    names = ["sig_exec_ok", "sig_nonempty", "sig_logprob", "sig_self_consistency", "sig_gliclass"]
    stacks = {}
    for f in (0, 1):
        d = df[df["qid"].map(qids_by_fold) != f]  # train on the other fold
        X = np.where(np.isfinite(d[names].to_numpy(float)), d[names].to_numpy(float), -5.0)
        sc = StandardScaler().fit(X)
        stacks[f] = (sc, LogisticRegression(max_iter=2000).fit(sc.transform(X), d["correct"].astype(int)), names)
    return stacks


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--cache", default="data")
    ap.add_argument("--signals", required=True)
    ap.add_argument("--configs", default="A,B,C,D,F,SC")
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--judge-model", default="zai-org/GLM-4.5-Air-FP8")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    qs = bird.load_questions(cache_dir=a.cache)
    qs = qs[: a.limit] if a.limit else qs
    rng = np.random.default_rng(0)
    fold = dict(zip([q.qid for q in qs], rng.permutation(len(qs)) % 2))
    stacks = fit_half_stacks(a.signals, fold)

    from vfd.schema_link import GLiClassLinker
    from vfd.signals import GLiClassScorer

    gli = GLiClassScorer(a.gliclass)
    judge = Chat(a.judge_model, base_url=os.environ.get("JUDGE_BASE_URL"), no_think=True) if os.environ.get("JUDGE_BASE_URL") else None
    schemas, paths = {}, {}
    for q in qs:
        if q.db_id not in schemas:
            paths[q.db_id] = bird.find_db(a.db_root, q.db_id)
            schemas[q.db_id] = bird.load_schema(paths[q.db_id])
    golds = {q.qid: bird.execute(paths[q.db_id], q.gold_sql, timeout_s=60) for q in qs}
    out = run / "harness_rows.jsonl"
    done = {(r["qid"], r["config"]) for r in map(json.loads, out.open())} if out.exists() else set()
    import threading

    lock = threading.Lock()  # GLiClass on one GPU: serialise its calls across worker threads

    class LockedGli:
        def scores(self, texts, labels):
            with lock:
                return gli.scores(texts, labels)

    class LockedLinker:
        def __init__(self):
            self.inner = GLiClassLinker(gli)

        def scores(self, query, labels):
            with lock:
                return self.inner.scores(query, labels)

    for cname in a.configs.split(","):
        cfg = CONFIGS[cname]
        if cfg.verifier == "cascade" and judge is None:
            print(f"[p2] skip {cname}: no JUDGE_BASE_URL")
            continue
        todo = [q for q in qs if (q.qid, cname) not in done]
        t0 = time.time()

        def one(q):
            tools = Tools(chat=Chat(a.model), judge=judge, gliclass=LockedGli(), linker=LockedLinker(),
                          stack=stacks[fold[q.qid]])
            try:
                r = answer(q, paths[q.db_id], schemas[q.db_id], cfg, tools)
                g = golds[q.qid]
                r["correct"] = bool(g.ok and r["result_key"] == g.key)
            except Exception as e:  # one failed question must not stop the sweep
                r = {"qid": q.qid, "config": cname, "error": f"{type(e).__name__}: {str(e)[:200]}", "correct": False}
            r["fold"] = int(fold[q.qid])
            return r

        rows = pmap(one, todo, workers=a.workers)
        with out.open("a") as f:
            for r in rows:
                f.write(json.dumps(r) + "\n")
        print(f"[p2] {cname}: {len(rows)} questions in {time.time() - t0:.0f}s, accuracy "
              f"{np.mean([r['correct'] for r in rows]):.3f}", flush=True)
    df = pd.DataFrame([json.loads(x) for x in out.open()])
    agg = df.groupby("config").agg(accuracy=("correct", "mean"), n=("correct", "size"),
                                   llm_calls=("llm_calls", "mean"), judge_calls=("judge_calls", "mean"),
                                   seconds=("seconds", "mean"), schema_chars=("schema_chars", "mean"),
                                   tokens=("completion_tokens", "mean"))
    # paired comparison to A: per-question difference, standard error
    base = df[df.config == "A"].set_index("qid")["correct"].astype(float)
    for c in agg.index:
        d = (df[df.config == c].set_index("qid")["correct"].astype(float) - base).dropna()
        agg.loc[c, "vs_A"] = d.mean()
        agg.loc[c, "vs_A_se"] = d.std(ddof=1) / np.sqrt(len(d)) if len(d) > 1 else np.nan
    md = ["# P2: SQL harness ablation on BIRD Arcwise-Plat", "",
          "A = LLM alone (greedy, full schema); B = + GLiClass schema linking; C = + GLiClass triage and one revision; "
          "D = + 8 candidates picked by a cheap verifier stack (fitted on the other half of the questions); "
          "F = D + LLM judge on the uncertain band; SC = 8 candidates, majority result (no small models).", "",
          agg.round(4).to_markdown()]
    (run / "results.md").write_text("\n".join(md) + "\n")
    json.dump(agg.round(5).to_dict("index"), open(run / "metrics.json", "w"), indent=1)
    print("\n".join(md))


if __name__ == "__main__":
    main()
