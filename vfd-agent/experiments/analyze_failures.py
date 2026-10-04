"""Why does the SQL harness fail? Classify every wrong answer of one configuration by comparing its result with gold.

  python experiments/analyze_failures.py --rows RUN/harness_rows.jsonl --config A --db-root data/bird
Categories (first match wins): exec_error, empty, extra_columns (gold result = a projection of the prediction),
missing_columns, wrong_row_count, wrong_values. Also checks every string literal compared with "=" in the predicted SQL
against the database ("literal not in DB" = a value-linking error a deterministic check could catch).
"""
from __future__ import annotations

import argparse
import itertools
import json
import re
import sqlite3
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd import bird  # noqa: E402

LIT = re.compile(r'([\w."`\[\]]+)\s*=\s*\'([^\']{1,80})\'')


def literal_misses(db: Path, sql: str) -> list[str]:
    out = []
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    try:
        tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")]
        for col, val in LIT.findall(sql):
            c = col.split(".")[-1].strip('"`[]')
            found = False
            for t in tables:
                cols = [r[1] for r in con.execute(f'PRAGMA table_info("{t}")')]
                if c not in cols:
                    continue
                try:
                    if con.execute(f'SELECT 1 FROM "{t}" WHERE "{c}" = ? LIMIT 1', (val,)).fetchone():
                        found = True
                        break
                except sqlite3.Error:
                    pass
            if not found:
                out.append(f"{c}='{val}'")
    finally:
        con.close()
    return out


def classify(pred: bird.ExecResult, gold: bird.ExecResult) -> str:
    if not pred.ok:
        return "exec_error"
    pr, gr = pred.rows or [], gold.rows or []
    if not pr and gr:
        return "empty"
    pc, gc = (len(pr[0]) if pr else 0), (len(gr[0]) if gr else 0)
    if pc > gc and gc:
        gset = {tuple(map(bird._norm, r)) for r in gr}
        for idx in itertools.combinations(range(pc), gc):
            if {tuple(bird._norm(r[i]) for i in idx) for r in pr} == gset:
                return "extra_columns"
        return "wrong_columns"
    if pc < gc:
        return "missing_columns"
    if len(set(pr)) != len(set(gr)):
        return "wrong_row_count"
    return "wrong_values"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", required=True)
    ap.add_argument("--config", default="A")
    ap.add_argument("--db-root", default="data/bird")
    ap.add_argument("--cache", default="data")
    a = ap.parse_args()
    qs = {q.qid: q for q in bird.load_questions(cache_dir=a.cache)}
    rows = [r for r in map(json.loads, open(a.rows)) if r["config"] == a.config]
    cats, lits, ex = Counter(), Counter(), {}
    for r in rows:
        q = qs[r["qid"]]
        db = bird.find_db(a.db_root, q.db_id)
        miss = literal_misses(db, r["sql"] or "")
        lits["wrong" if not r["correct"] else "right", bool(miss)] += 1
        if r["correct"]:
            continue
        c = classify(bird.execute(db, r["sql"] or "SELECT 1"), bird.execute(db, q.gold_sql, timeout_s=60))
        cats[c] += 1
        ex.setdefault(c, []).append((q.question[:120], q.evidence[:120], (r["sql"] or "")[:300], q.gold_sql[:300], miss))
    n = len(rows)
    print(f"config {a.config}: {n} questions, {n - sum(cats.values())} correct")
    for c, k in cats.most_common():
        print(f"  {c:16s} {k:4d}  ({k / n:.1%} of all)")
    print("literal not in DB (wrong answers):", lits["wrong", True], "| (right answers):", lits["right", True])
    for c in cats:
        print(f"\n== {c}")
        for e in ex[c][:3]:
            print(" Q:", e[0], "\n E:", e[1], "\n P:", e[2].replace("\n", " "), "\n G:", e[3].replace("\n", " "),
                  "\n missing literals:", e[4])


if __name__ == "__main__":
    main()
