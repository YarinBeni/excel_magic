"""P4: auto-research over the SQL harness configuration on BIRD Arcwise-Plat, with the guards in vfd.autoresearch.

The verifier stack is fitted only on the P0 candidates of SEARCH + ACCEPT questions; HELDOUT questions are touched
once, at the end. OPENAI_BASE_URL serves the generator; RESEARCHER_BASE_URL (optional) a researcher LLM, else the
generator proposes the experiments too.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import threading
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402
from vfd.autoresearch import AutoResearch  # noqa: E402
from vfd.harness import Config, Tools, answer  # noqa: E402
from vfd.llm import Chat, pmap  # noqa: E402

SPACE = {"link": [None, "gliclass", "rerank"], "link_k": [10, 20, 40], "triage": [False, True],
         "triage_threshold": [0.3, 0.5, 0.7], "n": [1, 4, 8, 16], "verifier": ["greedy", "self_consistency", "stack"]}
INIT = {"link": None, "link_k": 20, "triage": False, "triage_threshold": 0.5, "n": 1, "verifier": "greedy"}


def fit_stack(signals_csv, qids):
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler

    df = pd.read_csv(signals_csv)
    df = df[df["gold_ok"] & df["qid"].isin(set(qids))]
    names = ["sig_exec_ok", "sig_nonempty", "sig_logprob", "sig_self_consistency", "sig_gliclass"]
    X = np.where(np.isfinite(df[names].to_numpy(float)), df[names].to_numpy(float), -5.0)
    sc = StandardScaler().fit(X)
    return sc, LogisticRegression(max_iter=2000).fit(sc.transform(X), df["correct"].astype(int)), names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--cache", default="data")
    ap.add_argument("--signals", required=True)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--budget", type=int, default=20)
    ap.add_argument("--workers", type=int, default=16)
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    qs = {q.qid: q for q in bird.load_questions(cache_dir=a.cache)}
    paths, schemas = {}, {}
    for q in qs.values():
        if q.db_id not in schemas:
            paths[q.db_id] = bird.find_db(a.db_root, q.db_id)
            schemas[q.db_id] = bird.load_schema(paths[q.db_id])
    golds = {i: bird.execute(paths[q.db_id], q.gold_sql, timeout_s=60) for i, q in qs.items()}

    from vfd.schema_link import GLiClassLinker, RerankScorer
    from vfd.signals import GLiClassScorer

    gli = GLiClassScorer("knowledgator/gliclass-large-v3.0")
    rerank = RerankScorer()
    lock = threading.Lock()

    class Locked:
        def __init__(self, inner):
            self.inner = inner

        def scores(self, *x):
            with lock:
                return self.inner.scores(*x)

    linkers = {"gliclass": Locked(GLiClassLinker(gli)), "rerank": Locked(rerank)}
    researcher = Chat(a.model, base_url=os.environ.get("RESEARCHER_BASE_URL"))
    ar = AutoResearch(lambda c, q: {}, SPACE, researcher, budget=a.budget, z=1.0, seed=0)
    S, A, H = ar.split(sorted(qs))
    stack = fit_stack(a.signals, S + A)
    cache: dict = {}
    log = run / "autoresearch_evals.jsonl"

    def evaluate(cfg: dict, qids: list) -> dict:
        key = json.dumps(cfg, sort_keys=True)
        c = Config("auto", link=cfg["link"], link_k=cfg["link_k"], triage=cfg["triage"],
                   triage_threshold=cfg["triage_threshold"], n=cfg["n"], verifier=cfg["verifier"])
        todo = [q for q in qids if (key, q) not in cache]

        def one(qid):
            q = qs[qid]
            tools = Tools(chat=Chat(a.model), gliclass=Locked(gli), linker=linkers.get(cfg["link"]), stack=stack)
            try:
                r = answer(q, paths[q.db_id], schemas[q.db_id], c, tools)
                return qid, float(golds[qid].ok and r["result_key"] == golds[qid].key)
            except Exception:
                return qid, 0.0

        for qid, s in pmap(one, todo, workers=a.workers):
            cache[(key, qid)] = s
        out = {q: cache[(key, q)] for q in qids}
        with log.open("a") as f:
            f.write(json.dumps({"config": cfg, "n": len(qids), "mean": float(np.mean(list(out.values())))}) + "\n")
        return out

    ar.evaluate = evaluate
    res = ar.run(INIT, sorted(qs))
    json.dump(res, open(run / "metrics.json", "w"), indent=1, default=str)
    md = ["# P4: guarded auto-research over the SQL harness (BIRD Arcwise-Plat)", "",
          f"Search {res['n_search']} / accept {res['n_accept']} / held-out {res['n_heldout']} questions; "
          f"{res['attempts']} attempts, {res['kept']} kept.", "",
          f"Initial {json.dumps(res['initial'])}", f"Final   {json.dumps(res['final'])}", "",
          f"Held-out accuracy: initial {res['heldout_initial']:.4f} -> final {res['heldout_final']:.4f} "
          f"(paired delta {res['heldout_delta']:+.4f}, se {res['heldout_se']:.4f})", "",
          "| # | knob | value | hypothesis | search delta | se | accept delta | kept |", "|---|---|---|---|---|---|---|---|"]
    md += [f"| {h['i']} | {h['knob']} | {h['value']} | {h['hypothesis'][:60]} | {h['search_delta']:+.4f} | "
           f"{h['search_se']:.4f} | {h['accept_delta']:+.4f} | {'yes' if h['kept'] else 'no'} |" for h in res["history"]]
    (run / "results.md").write_text("\n".join(md) + "\n")
    print("\n".join(md))


if __name__ == "__main__":
    main()
