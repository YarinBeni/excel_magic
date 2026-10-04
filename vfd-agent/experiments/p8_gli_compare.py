"""Q5b: zero-shot head-to-head of open GLi-family models on the two jobs GLiClass had in the harness.

  checker  AUROC of P(correct) over all P0 SQL candidates (question, evidence, SQL, result preview)
  linker   recall@k of the gold SQL's columns (corrected descriptions), all questions
  python experiments/p8_gli_compare.py --run DIR --p0 P0RUN --desc-root R --cache C --models m1,m2,...
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
from vfd.schema_link import GLiClassLinker, column_label, gold_columns  # noqa: E402
from vfd.signals import candidate_text  # noqa: E402

POS, NEG = "the query result correctly answers the question", "the query result does not answer the question"


def scorer(mid: str):
    if "gliclass" in mid.lower():
        from vfd.signals import GLiClassScorer
        return GLiClassScorer(mid)
    from vfd.gliner2_scorer import GLiNER2Scorer
    return GLiNER2Scorer(mid)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--p0", required=True)
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--desc-root", default="")
    ap.add_argument("--cache", default="data")
    ap.add_argument("--models", default="knowledgator/gliclass-large-v3.0,fastino/gliner2-large-v1,"
                                        "fastino/gliner2.5-base-v1,fastino/GLiNER2.5-Decide,fastino/GLiNER2.5-Decide-1B")
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    from sklearn.metrics import roc_auc_score

    qs = {q.qid: q for q in bird.load_questions(cache_dir=a.cache)}
    c = pd.DataFrame([json.loads(x) for x in (Path(a.p0) / "candidates.jsonl").open()])
    c = c[c.gold_ok].reset_index(drop=True)
    texts = [candidate_text(r, {"question": qs[r["qid"]].question, "evidence": qs[r["qid"]].evidence})
             for r in c.to_dict("records")]
    schemas, labels = {}, {}
    for q in qs.values():
        if q.db_id not in schemas:
            p = bird.find_db(a.db_root, q.db_id)
            schemas[q.db_id] = bird.load_schema(p, desc_dir=Path(a.desc_root) / q.db_id if a.desc_root else None)
            labels[q.db_id] = {(x.table.lower(), x.name.lower()): column_label(x) for x in schemas[q.db_id].columns}
    res = {}
    for mid in a.models.split(","):
        r = {}
        try:
            s = scorer(mid)
            t0 = time.time()
            sc = s.scores(texts, [POS, NEG])
            r["checker_ms"] = 1000 * (time.time() - t0) / len(texts)
            r["checker_auroc_pos"] = float(roc_auc_score(c.correct.astype(int), sc[:, 0]))
            r["checker_auroc_diff"] = float(roc_auc_score(c.correct.astype(int), sc[:, 0] - sc[:, 1]))
            lk = GLiClassLinker(s)
            rec = {k: [] for k in (5, 10, 20)}
            t0 = time.time()
            for q in qs.values():
                gold = gold_columns(q.gold_sql, schemas[q.db_id])
                if not gold:
                    continue
                keys = list(labels[q.db_id])
                o = [keys[j] for j in np.argsort(-lk.scores(f"{q.question} {q.evidence}", [labels[q.db_id][k] for k in keys]))]
                for k in rec:
                    rec[k].append(len(gold & set(o[:k])) / len(gold))
            r["linker_ms_per_question"] = 1000 * (time.time() - t0) / max(1, len(rec[5]))
            r.update({f"linker_recall@{k}": float(np.mean(v)) for k, v in rec.items()})
            del s
            import torch
            torch.cuda.empty_cache()
        except Exception as e:
            r["error"] = f"{type(e).__name__}: {str(e)[:300]}"
        res[mid] = r
        print(mid, json.dumps(r), flush=True)
    json.dump(res, open(run / "metrics.json", "w"), indent=1)
    df = pd.DataFrame(res).T
    (run / "results.md").write_text("# Q5b: zero-shot GLi-family models, BIRD checker and linker\n\n"
                                    f"{len(texts)} candidates; {len(qs)} questions.\n\n" + df.round(4).to_markdown() + "\n")
    print(df.round(4).to_markdown())


if __name__ == "__main__":
    main()
