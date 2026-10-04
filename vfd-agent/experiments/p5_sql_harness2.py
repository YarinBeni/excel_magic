"""Q1/Q2: harness v2 ladder on BIRD Arcwise-Plat-Full. The question split is the one P4 used (seed 0): DEV = search +
accept (373 questions) for every design decision, HELDOUT (125) scored once at the end.

  python experiments/p5_sql_harness2.py --run DIR --split dev --configs A,K1,K2,K3,G,GD,SC8,G8 --model M --desc-root D
OPENAI_BASE_URL serves the model. --desc-root holds <db_id>/database_description/*.csv (corrected descriptions).
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import threading
import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402
from vfd.autoresearch import AutoResearch  # noqa: E402
from vfd.harness2 import CONFIGS2, answer2, profile_schema  # noqa: E402
from vfd.llm import Chat, pmap  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--desc-root", default="")
    ap.add_argument("--cache", default="data")
    ap.add_argument("--split", default="dev", choices=["dev", "heldout", "all"])
    ap.add_argument("--configs", default="A,K1,K2,K3,G")
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--tag", default="", help="label for the model in the output rows")
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    qs = {q.qid: q for q in bird.load_questions(cache_dir=a.cache)}
    S, A_, H = AutoResearch(lambda c, q: {}, {}, None, seed=0).split(sorted(qs))
    ids = {"dev": S + A_, "heldout": H, "all": sorted(qs)}[a.split]
    paths, plain, desc, prof = {}, {}, {}, {}
    for i in ids:
        d = qs[i].db_id
        if d in paths:
            continue
        paths[d] = bird.find_db(a.db_root, d)
        plain[d] = bird.load_schema(paths[d], desc_dir=Path("/nonexistent"))
        desc[d] = bird.load_schema(paths[d], desc_dir=Path(a.desc_root) / d if a.desc_root else None)
        prof[d] = profile_schema(paths[d], copy.deepcopy(desc[d]))
    golds = {i: bird.execute(paths[qs[i].db_id], qs[i].gold_sql, timeout_s=60) for i in ids}
    gli = None
    if any(CONFIGS2[c].doubt for c in a.configs.split(",")):
        from vfd.signals import GLiClassScorer

        inner, lock = GLiClassScorer(a.gliclass), threading.Lock()

        class Locked:
            def scores(self, texts, labels):
                with lock:
                    return inner.scores(texts, labels)

        gli = Locked()
    tag = a.tag or a.model.split("/")[-1]
    out = run / "harness2_rows.jsonl"
    done = {(r["qid"], r["config"], r["model"]) for r in map(json.loads, out.open())} if out.exists() else set()
    for cname in a.configs.split(","):
        cfg = CONFIGS2[cname]
        todo = [i for i in ids if (i, cname, tag) not in done]
        t0 = time.time()

        def one(i):
            q = qs[i]
            sch = prof[q.db_id] if cfg.profile else desc[q.db_id] if cfg.desc else plain[q.db_id]
            try:
                r = answer2(q, paths[q.db_id], bird.render_schema(sch), cfg, Chat(a.model), gli)
                r["correct"] = bool(golds[i].ok and r["result_key"] == golds[i].key)
            except Exception as e:
                r = {"qid": i, "config": cname, "error": f"{type(e).__name__}: {str(e)[:200]}", "correct": False}
            r.update(model=tag, split=a.split)
            return r

        rows = pmap(one, todo, workers=a.workers)
        with out.open("a") as f:
            for r in rows:
                f.write(json.dumps(r) + "\n")
        print(f"[h2] {tag} {cname}: {len(rows)} q in {time.time() - t0:.0f}s, acc "
              f"{np.mean([r['correct'] for r in rows]) if rows else float('nan'):.3f}", flush=True)
    report(run)


def report(run: Path) -> None:
    df = pd.DataFrame([json.loads(x) for x in (run / "harness2_rows.jsonl").open()])
    lines = ["# Harness v2 ladder (BIRD Arcwise-Plat-Full)", ""]
    for (model, split), g in df.groupby(["model", "split"]):
        base = g[g.config == "A"].set_index("qid")["correct"].astype(float)
        rows = []
        for c, gc in g.groupby("config", sort=False):
            d = (gc.set_index("qid")["correct"].astype(float) - base).dropna()
            rows.append({"config": c, "n": len(gc), "accuracy": gc.correct.mean(),
                         "vs_A": d.mean() if len(d) else np.nan,
                         "se": d.std(ddof=1) / np.sqrt(len(d)) if len(d) > 1 else np.nan,
                         "llm_calls": gc.get("llm_calls", pd.Series(dtype=float)).mean(),
                         "revisions": gc.get("revisions", pd.Series(dtype=float)).mean(),
                         "schema_chars": gc.get("schema_chars", pd.Series(dtype=float)).mean()})
        lines += [f"## {model}, {split}", "", pd.DataFrame(rows).round(4).to_markdown(index=False), ""]
    (run / "results.md").write_text("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
