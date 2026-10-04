"""Q5: fine-tune GLiClass (the LLM stays zero-shot) on the DEV questions only, test on the HELDOUT questions.

  checker  (question, evidence, SQL, result preview) -> "correct" / "wrong". Training rows: the P0 candidates (8+1 per
           question, labelled by execution against gold) of DEV questions. Test: AUROC on HELDOUT candidates, next to
           zero-shot GLiClass, self-consistency and the LLM judge on the same rows (columns of P0 signals.csv).
  linker   (question + evidence) -> which schema columns the gold SQL uses (multi-label over the DB's column labels,
           with corrected descriptions). Test: recall@k of gold columns on HELDOUT questions, zero-shot vs tuned.

  python experiments/p7_gliclass_finetune.py --run DIR --p0 P0RUN --db-root D --desc-root R --cache C
DEV / HELDOUT = the P4 split (seed 0): search + accept (373) / held-out (125).
"""
from __future__ import annotations

import argparse
import json
import os
import random
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402
from vfd.autoresearch import AutoResearch  # noqa: E402
from vfd.schema_link import GLiClassLinker, column_label, gold_columns  # noqa: E402
from vfd.signals import GLiClassScorer, candidate_text  # noqa: E402

BASE = "knowledgator/gliclass-large-v3.0"
POS, NEG = "the query result correctly answers the question", "the query result does not answer the question"


def auc(y, s):
    from sklearn.metrics import roc_auc_score
    return float(roc_auc_score(np.asarray(y, int), np.asarray(s, float)))


def train(examples: list[dict], out_dir: Path, epochs: int, lr: float, bs: int, max_length: int) -> Path:
    import torch
    from gliclass import GLiClassModel
    from gliclass.data_processing import AugmentationConfig, DataCollatorWithPadding, GLiClassDataset
    from gliclass.training import Trainer, TrainingArguments
    from transformers import AutoTokenizer

    model = GLiClassModel.from_pretrained(BASE)
    tok = AutoTokenizer.from_pretrained(BASE, add_prefix_space=True)
    ds = GLiClassDataset(examples, tok, AugmentationConfig(enabled=False), max_length=max_length,
                         problem_type="multi_label_classification", architecture_type=model.config.architecture_type,
                         prompt_first=model.config.prompt_first, shuffle_labels=True)
    args = TrainingArguments(output_dir=str(out_dir / "ckpt"), learning_rate=lr, others_lr=lr * 3, weight_decay=0.01,
                             others_weight_decay=0.01, per_device_train_batch_size=bs, num_train_epochs=epochs,
                             warmup_ratio=0.05, lr_scheduler_type="linear", save_strategy="no", logging_steps=50,
                             report_to="none", bf16=torch.cuda.is_available(), dataloader_num_workers=0, dataloader_pin_memory=False,
                             remove_unused_columns=False)
    Trainer(model=model, args=args, train_dataset=ds,
            data_collator=DataCollatorWithPadding(device="cuda:0" if torch.cuda.is_available() else "cpu")).train()
    model.save_pretrained(out_dir)
    tok.save_pretrained(out_dir)
    return out_dir


def checker(a, qs, dev, held, run: Path, res: dict):
    p0 = Path(a.p0)
    cands = pd.DataFrame([json.loads(x) for x in (p0 / "candidates.jsonl").open()])
    sig = pd.read_csv(p0 / "signals.csv")
    cands = cands[cands.gold_ok & cands.qid.isin(qs)].reset_index(drop=True)
    qd = {i: {"question": q.question, "evidence": q.evidence} for i, q in qs.items()}
    cands["text"] = [candidate_text(r, qd[r["qid"]]) for r in cands.to_dict("records")]
    tr = cands[cands.qid.isin(dev)]
    ex = [{"text": t, "all_labels": [POS, NEG], "true_labels": [POS if c else NEG]} for t, c in zip(tr.text, tr.correct)]
    random.Random(0).shuffle(ex)
    path = train(ex, Path(a.models) / "checker", a.epochs, a.lr, a.bs, a.max_length)
    te = cands[cands.qid.isin(held)].reset_index(drop=True)
    tuned = GLiClassScorer(str(path)).scores(te.text.tolist(), [POS, NEG])
    te["tuned"] = tuned[:, 0] - tuned[:, 1]
    te = te.merge(sig[["qid", "cand", "sig_gliclass", "sig_judge", "sig_self_consistency", "sig_logprob"]],
                  on=["qid", "cand"], how="left")
    res["checker"] = {"n_train": len(ex), "train_pos": float(tr.correct.mean()), "n_test": len(te),
                      "auroc": {c: auc(te.correct, te[c].fillna(te[c].min())) for c in
                                ["tuned", "sig_gliclass", "sig_judge", "sig_self_consistency", "sig_logprob"]}}
    # best-of-candidates accuracy on held-out questions: pick max score per question
    for c in ["tuned", "sig_gliclass", "sig_judge", "sig_self_consistency"]:
        pick = te.sort_values(c, ascending=False).groupby("qid").head(1)
        res["checker"].setdefault("best_of_n", {})[c] = float(pick.correct.mean())
    res["checker"]["best_of_n"]["greedy"] = float(te[te.greedy].correct.mean())
    res["checker"]["best_of_n"]["oracle"] = float(te.groupby("qid").correct.max().mean())
    te[["qid", "cand", "correct", "tuned"]].to_csv(run / "checker_scores.csv", index=False)


def linker(a, qs, dev, held, res: dict):
    schemas, labels = {}, {}
    for q in qs.values():
        if q.db_id not in schemas:
            p = bird.find_db(a.db_root, q.db_id)
            schemas[q.db_id] = bird.load_schema(p, desc_dir=Path(a.desc_root) / q.db_id if a.desc_root else None)
            labels[q.db_id] = {(c.table.lower(), c.name.lower()): column_label(c) for c in schemas[q.db_id].columns}
    rng = random.Random(0)
    ex = []
    for i in dev:
        q = qs[i]
        gold = gold_columns(q.gold_sql, schemas[q.db_id])
        lab = labels[q.db_id]
        pos = [lab[g] for g in gold if g in lab]
        neg = [v for k, v in lab.items() if k not in gold]
        if not pos:
            continue
        for _ in range(a.link_views):  # several 40-label views of the same schema
            sample = pos + rng.sample(neg, min(len(neg), 40 - len(pos)))
            rng.shuffle(sample)
            ex.append({"text": f"{q.question} {q.evidence}", "all_labels": sample, "true_labels": pos})
    path = train(ex, Path(a.models) / "linker", a.epochs, a.lr, a.bs, a.max_length)
    out = {}
    for name, mid in [("zero_shot", BASE), ("tuned", str(path))]:
        lk = GLiClassLinker(GLiClassScorer(mid))
        rec = {k: [] for k in (5, 10, 20)}
        allk = {k: [] for k in (10, 20)}
        for i in held:
            q = qs[i]
            gold = gold_columns(q.gold_sql, schemas[q.db_id])
            keys = list(labels[q.db_id])
            if not gold:
                continue
            s = lk.scores(f"{q.question} {q.evidence}", [labels[q.db_id][k] for k in keys])
            order = [keys[j] for j in np.argsort(-s)]
            for k in rec:
                rec[k].append(len(gold & set(order[:k])) / len(gold))
            for k in allk:
                allk[k].append(float(gold <= set(order[:k])))
        out[name] = {**{f"recall@{k}": float(np.mean(v)) for k, v in rec.items()},
                     **{f"all@{k}": float(np.mean(v)) for k, v in allk.items()}, "n": len(rec[5])}
    res["linker"] = {"n_train_examples": len(ex), **out}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--p0", required=True, help="P0 run dir with candidates.jsonl and signals.csv")
    ap.add_argument("--db-root", default=os.environ.get("BIRD_ROOT", "data/bird"))
    ap.add_argument("--desc-root", default="")
    ap.add_argument("--cache", default="data")
    ap.add_argument("--models", default="artifacts/gliclass_ft")
    ap.add_argument("--tasks", default="checker,linker")
    ap.add_argument("--epochs", type=int, default=3)
    ap.add_argument("--lr", type=float, default=1e-5)
    ap.add_argument("--bs", type=int, default=8)
    ap.add_argument("--max-length", type=int, default=1024)
    ap.add_argument("--link-views", type=int, default=3)
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    qs = {q.qid: q for q in bird.load_questions(cache_dir=a.cache)}
    S, A_, H = AutoResearch(lambda c, q: {}, {}, None, seed=0).split(sorted(qs))
    dev, held = S + A_, H
    res: dict = {"n_dev": len(dev), "n_heldout": len(held)}
    if "checker" in a.tasks:
        checker(a, qs, set(dev), set(held), run, res)
    if "linker" in a.tasks:
        linker(a, qs, dev, held, res)
    json.dump(res, open(run / "metrics.json", "w"), indent=1)
    md = ["# Q5: GLiClass fine-tuned on DEV questions (LLM unchanged), tested on HELDOUT", "", "```",
          json.dumps(res, indent=1), "```"]
    (run / "results.md").write_text("\n".join(md) + "\n")
    print("\n".join(md))


if __name__ == "__main__":
    main()
