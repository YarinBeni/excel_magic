"""P0: verifier study on BIRD Arcwise-Plat (498 fully corrected questions).

  python experiments/p0_verifier_study.py generate --db-root $BIRD --run DIR --model Qwen/Qwen3-Coder-30B-A3B-Instruct
  python experiments/p0_verifier_study.py judge    --db-root $BIRD --run DIR --model zai-org/GLM-4.5-Air-FP8
  python experiments/p0_verifier_study.py encoders --db-root $BIRD --run DIR          (GPU: GLiClass, NLI)
  python experiments/p0_verifier_study.py link     --db-root $BIRD --run DIR          (GPU: schema-linking scorers)
  python experiments/p0_verifier_study.py report   --run DIR

Each stage reads and writes files in the run directory, so stages can run in separate processes (one vLLM server at a
time on one GPU) and resume after a failure.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402


def load_q(a) -> list[bird.Question]:
    qs = bird.load_questions(a.questions, cache_dir=a.cache)
    return qs[: a.limit] if a.limit else qs


def qdict(qs) -> dict[int, dict]:
    return {q.qid: {"question": q.question, "evidence": q.evidence, "db_id": q.db_id} for q in qs}


def cmd_generate(a):
    from vfd.candidates import generate
    from vfd.llm import Chat

    qs = load_q(a)
    chat = Chat(a.model)
    t0 = time.time()
    rows = generate(qs, a.db_root, chat, Path(a.run) / "candidates.jsonl", n_samples=a.n, workers=a.workers)
    df = pd.DataFrame(rows)
    print(f"[generate] {len(df)} candidates for {df.qid.nunique()} questions in {time.time() - t0:.0f}s; "
          f"greedy acc {df[df.greedy].correct.mean():.3f}; oracle {df.groupby('qid').correct.any().mean():.3f}; "
          f"gold executes {df.groupby('qid').gold_ok.first().mean():.3f}")


def candidates(a) -> pd.DataFrame:
    return pd.DataFrame([json.loads(x) for x in (Path(a.run) / "candidates.jsonl").open()])


def cmd_judge(a):
    from vfd.llm import Chat
    from vfd.signals import judge_signal

    qs = load_q(a)
    df = candidates(a)
    schemas = {}
    for db in df.db_id.unique():
        schemas[db] = bird.render_schema(bird.load_schema(bird.find_db(a.db_root, db)))
    chat = Chat(a.model, no_think=True)
    # single-call latency on 20 items, then the parallel run
    from vfd.signals import JUDGE_SYSTEM, candidate_text  # noqa: F401
    t0 = time.time()
    sub = df.head(20)
    judge_signal(sub, qdict(qs), schemas, chat, workers=1)
    single = (time.time() - t0) / len(sub)
    out, lat = judge_signal(df, qdict(qs), schemas, chat, workers=a.workers)
    lat["judge_s_per_call_sequential"] = single
    out[["qid", "cand", "sig_judge"]].to_csv(Path(a.run) / "judge.csv", index=False)
    json.dump(lat, open(Path(a.run) / "latency_judge.json", "w"), indent=1)
    print("[judge] done", lat)


def cmd_encoders(a):
    from vfd.signals import GLiClassScorer, NLIScorer, encoder_signals

    qs = load_q(a)
    df = candidates(a)
    gli = GLiClassScorer(a.gliclass, device="cuda:0")
    nli = NLIScorer(a.nli, device=0)
    out, lat = encoder_signals(df, qdict(qs), gli, nli)
    keep = ["qid", "cand"] + [c for c in out.columns if c.startswith(("sig_gliclass", "sig_nli", "tax_"))]
    out[keep].to_csv(Path(a.run) / "encoders.csv", index=False)
    json.dump(lat, open(Path(a.run) / "latency_encoders.json", "w"), indent=1)
    print("[encoders] done", lat)


def cmd_link(a):
    from vfd.schema_link import (EmbedScorer, GLiClassLinker, GLiNERLinker, RerankScorer, evaluate_linking,
                                 gold_columns, lexical_scores)
    from vfd.signals import GLiClassScorer

    qs = load_q(a)
    schemas = {}
    items = []
    for q in qs:
        if q.db_id not in schemas:
            schemas[q.db_id] = bird.load_schema(bird.find_db(a.db_root, q.db_id))
        gold = gold_columns(q.gold_sql, schemas[q.db_id])
        if gold:
            items.append((q.qid, f"{q.question} {q.evidence}", schemas[q.db_id], gold))
    scorers = {"lexical": lexical_scores, "embed_bge_small": EmbedScorer(), "rerank_bge_m3": RerankScorer(),
               "gliclass_large_v3": GLiClassLinker(GLiClassScorer(a.gliclass)), "gliner_bi_base_v2": GLiNERLinker()}
    rows, lat = evaluate_linking(items, scorers)
    pd.DataFrame(rows).to_csv(Path(a.run) / "schema_link.csv", index=False)
    json.dump(lat, open(Path(a.run) / "latency_link.json", "w"), indent=1)
    print("[link] done", pd.DataFrame(rows).groupby("scorer")[["recall@5", "recall@10", "recall@20", "all@20"]].mean().round(3))


def cmd_report(a):
    from vfd.evaluate import summarize
    from vfd.signals import cheap_signals

    run = Path(a.run)
    df = cheap_signals(candidates(a))
    for f in ("encoders.csv", "judge.csv"):
        if (run / f).exists():
            df = df.merge(pd.read_csv(run / f), on=["qid", "cand"], how="left")
    res = summarize(df)
    frame = res.pop("_frame")
    lat = {}
    for f in run.glob("latency_*.json"):
        lat.update(json.load(open(f)))
    res["latency"] = lat
    link = None
    if (run / "schema_link.csv").exists():
        link = pd.read_csv(run / "schema_link.csv").groupby("scorer")[
            ["recall@5", "recall@10", "recall@20", "all@10", "all@20"]].mean()
        res["schema_link"] = link.round(4).to_dict("index")
    names = {"sig_exec_ok": "query executes", "sig_nonempty": "non-empty result", "sig_logprob": "generator logprob",
             "sig_self_consistency": "self-consistency", "sig_gliclass": "GLiClass large v3 (zero-shot)",
             "sig_nli": "NLI cross-encoder (DeBERTa-v3-large)", "sig_judge": "LLM judge (2nd family)",
             "stack_cheap": "stack: no-model signals", "stack_cheap_gli": "stack: + GLiClass (+ taxonomy)",
             "stack_cheap_gli_nli": "stack: + GLiClass + NLI", "stack_all": "stack: + judge"}
    md = [f"# P0 verifier study: {res['n_questions']} BIRD Arcwise-Plat questions, {res['n_candidates']} SQL candidates", "",
          f"Greedy accuracy {res['greedy_accuracy']:.3f}; oracle (any candidate correct) {res['oracle_accuracy']:.3f}; "
          f"{res['candidates_correct_share']:.3f} of candidates correct.", "",
          "| signal | AUROC | within-question AUROC | best-of-N accuracy | ECE raw | ECE isotonic |", "|---|---|---|---|---|---|"]
    for k, r in res["signals"].items():
        md.append(f"| {names.get(k, k)} | {r['auroc']:.3f} | {r['within_question_auroc']:.3f} | {r['best_of_n']:.3f} | "
                  f"{r.get('ece_raw', float('nan')):.3f} | {r['ece_isotonic']:.3f} |")
    if "cascade" in res:
        md += ["", "Cascade (stack with GLiClass; the most uncertain share goes to the judge):", "",
               "| judge share | AUROC | best-of-N |", "|---|---|---|"]
        md += [f"| {c['judge_share']:.0%} | {c['auroc']:.3f} | {c['best_of_n']:.3f} |" for c in res["cascade"]]
    if link is not None:
        md += ["", "Schema linking (recall of the gold SQL's columns among the top-k ranked columns):", "",
               link.round(3).to_markdown()]
    if lat:
        md += ["", "Latency (seconds per item): " + ", ".join(f"{k} {v:.4f}" for k, v in lat.items() if isinstance(v, float))]
    (run / "results.md").write_text("\n".join(md) + "\n")
    frame.drop(columns=["sql", "preview", "error"], errors="ignore").to_csv(run / "signals.csv", index=False)
    json.dump(res, open(run / "metrics.json", "w"), indent=1, default=float)
    print("\n".join(md))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("stage", choices=["generate", "judge", "encoders", "link", "report"])
    ap.add_argument("--run", required=True)
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--questions", default=None)
    ap.add_argument("--cache", default="data")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--n", type=int, default=8)
    ap.add_argument("--workers", type=int, default=32)
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    ap.add_argument("--nli", default="MoritzLaurer/deberta-v3-large-zeroshot-v2.0")
    a = ap.parse_args()
    Path(a.run).mkdir(parents=True, exist_ok=True)
    {"generate": cmd_generate, "judge": cmd_judge, "encoders": cmd_encoders, "link": cmd_link, "report": cmd_report}[a.stage](a)


if __name__ == "__main__":
    main()
