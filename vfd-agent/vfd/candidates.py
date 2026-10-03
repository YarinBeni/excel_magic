"""Stage A of the verifier study: for each question, one greedy SQL plus n sampled SQL from the generator LLM, each
executed and labelled correct / wrong against the corrected gold result."""
from __future__ import annotations

import json
from pathlib import Path

from .bird import Question, Schema, execute, find_db, load_schema, preview, render_schema, same_result
from .llm import Chat, extract_sql, pmap

SYSTEM = ("You are an expert SQLite analyst. Write one SQLite query that answers the question. Use only the tables and "
          "columns in the schema. Use the evidence to interpret terms. Return the query in a ```sql block and nothing else.")


def prompt(q: Question, schema_text: str) -> list[dict]:
    return [{"role": "system", "content": SYSTEM},
            {"role": "user", "content": f"Schema:\n{schema_text}\n\nEvidence: {q.evidence or '(none)'}\nQuestion: {q.question}"}]


def generate(questions: list[Question], db_root: str, chat: Chat, out_path: Path, n_samples: int = 8,
             temperature: float = 0.8, workers: int = 32) -> list[dict]:
    """Writes candidates.jsonl: one row per (question, candidate). Resumable: questions already in the file are skipped."""
    done = set()
    if out_path.exists():
        done = {json.loads(line)["qid"] for line in out_path.open()}
    schemas: dict[str, tuple[Path, Schema, str]] = {}
    for q in questions:
        if q.db_id not in schemas:
            p = find_db(db_root, q.db_id)
            s = load_schema(p)
            schemas[q.db_id] = (p, s, render_schema(s))
    todo = [q for q in questions if q.qid not in done]

    def one(q: Question) -> list[dict]:
        p, _, st = schemas[q.db_id]
        msgs = prompt(q, st)
        gold = execute(p, q.gold_sql, timeout_s=60)
        greedy = chat.samples(msgs, n=1, temperature=0.0)
        sampled = chat.samples(msgs, n=n_samples, temperature=temperature) if n_samples else []
        rows = []
        for k, smp in enumerate(greedy + sampled):
            sql = extract_sql(smp.text)
            r = execute(p, sql)
            rows.append({"qid": q.qid, "db_id": q.db_id, "cand": k, "greedy": k == 0, "sql": sql,
                         "exec_ok": r.ok, "n_rows": len(r.rows) if r.ok else -1, "result_key": r.key,
                         "preview": preview(r), "error": r.error, "exec_s": round(r.seconds, 3),
                         "mean_logprob": smp.mean_logprob, "sum_logprob": smp.sum_logprob, "n_tokens": smp.n_tokens,
                         "gold_ok": gold.ok, "correct": bool(same_result(r, gold))})
        return rows

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("a") as f:
        for chunk_start in range(0, len(todo), workers * 2):
            for rows in pmap(one, todo[chunk_start:chunk_start + workers * 2], workers=workers):
                for r in rows:
                    f.write(json.dumps(r) + "\n")
            f.flush()
    return [json.loads(line) for line in out_path.open()]
