"""SQL harness v2: the harness does the deterministic work, the LLM only writes and fixes SQL.

What changed against harness.py (each part is switchable, so every step is one ablation):
  context   K1 column descriptions (corrected BIRD description files)
            K2 a data profile per column (distinct count, NULL share, range, most frequent values), computed once per DB
            K3 a short rule card (general SQLite / question-reading rules; written before looking at any failure)
  gates     deterministic checks after execution: the query fails, a compared literal is not stored in that column
            (the harness looks up the closest stored values), the result is empty, or only NULL. A failing candidate
            goes back to the LLM with the check's message, up to `max_rev` rounds.
  doubt     optional online signal: GLiClass scores "the result answers the question"; below a threshold the LLM gets
            one more round with a neutral "re-read the question" message. The revision is kept only if it passes gates.
  select    n candidates; majority of result sets among candidates that pass all gates; ties go to the greedy one.
"""
from __future__ import annotations

import difflib
import re
import sqlite3
import time
from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

from . import bird
from .llm import extract_sql

SYSTEM = ("You are an expert SQLite analyst. Write one SQLite query that answers the question. Use only the tables and "
          "columns in the schema. Use the evidence to interpret terms. Return the query in a ```sql block and nothing else.")

RULES = """Rules:
1. Return exactly the columns the question asks for, in the order asked. Do not add ids, counts or helper columns.
2. Apply the evidence literally: its formulas, value mappings and thresholds are the definitions to use.
3. Compare with values exactly as they are stored (see the example values: case, spelling, format, units).
4. "Highest / lowest / most / least" means ORDER BY ... LIMIT 1 unless the question asks for all ties.
5. For ratios and percentages CAST to REAL before dividing; multiply by 100 only for a percentage.
6. Use DISTINCT when the question asks for different / unique entities or when a join can repeat rows.
7. Join only through the key columns shown; a wrong join silently changes counts.
8. Exclude NULLs from a column you sort or rank by."""

DOUBT_LABEL = "the query result correctly answers the question"


@dataclass
class Config2:
    name: str
    desc: bool = False
    profile: bool = False
    rules: bool = False
    gates: bool = False
    max_rev: int = 2
    doubt: bool = False
    doubt_threshold: float = 0.5
    n: int = 1


CONFIGS2 = {
    "A": Config2("A"),
    "K1": Config2("K1", desc=True),
    "K2": Config2("K2", desc=True, profile=True),
    "K3": Config2("K3", desc=True, profile=True, rules=True),
    "G": Config2("G", desc=True, profile=True, rules=True, gates=True),
    "GD": Config2("GD", desc=True, profile=True, rules=True, gates=True, doubt=True),
    "SC8": Config2("SC8", desc=True, profile=True, rules=True, n=8),
    "G8": Config2("G8", desc=True, profile=True, rules=True, gates=True, n=8),
}


# ----------------------------------------------------------------------------------------------- data profile
def profile_schema(db_path: Path, schema: bird.Schema, top: int = 5, max_distinct_listed: int = 30) -> bird.Schema:
    """Replace each column's example values with a profile: distinct count, NULL share, range or frequent values."""
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode(errors="replace")
    for t, cols in schema.tables.items():
        try:
            n = con.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
        except sqlite3.Error:
            continue
        for c in cols:
            q = f'"{c.name}"'
            try:
                nd, nn = con.execute(f'SELECT COUNT(DISTINCT {q}), COUNT({q}) FROM "{t}"').fetchone()
                parts = [f"{nd} distinct"]
                if n and nn < n:
                    parts.append(f"{100 * (n - nn) / n:.0f}% NULL")
                kind = con.execute(f'SELECT typeof({q}) FROM "{t}" WHERE {q} IS NOT NULL LIMIT 1').fetchone()
                if kind and kind[0] in ("integer", "real") and nd > max_distinct_listed:
                    lo, hi = con.execute(f'SELECT MIN({q}), MAX({q}) FROM "{t}"').fetchone()
                    parts.append(f"range {lo} to {hi}")
                else:
                    vals = con.execute(f'SELECT {q}, COUNT(*) c FROM "{t}" WHERE {q} IS NOT NULL GROUP BY {q} '
                                       f'ORDER BY c DESC LIMIT {top}').fetchall()
                    parts.append("values " + ", ".join(repr(str(v)[:30]) for v, _ in vals)
                                 + (" ..." if nd > top else ""))
                vd = c.values.split(" (", 1)[1].rstrip(")") if " (" in c.values else ""
                c.values = "; ".join(parts) + (f" ({vd})" if vd else "")
            except sqlite3.Error:
                pass
    con.close()
    return schema


# ----------------------------------------------------------------------------------------------- literal grounding
_LIT = re.compile(r'([\w."`\[\]]+)\s*(=|!=|<>|\bLIKE\b)\s*\'([^\']{1,80})\'', re.I)


@lru_cache(maxsize=4096)
def _stored_values(db: str, column: str) -> tuple[tuple[str, str], ...]:
    """(table, value) pairs of up to 20k distinct values of every column with this name."""
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode(errors="replace")
    out = []
    try:
        for (t,) in con.execute("SELECT name FROM sqlite_master WHERE type='table'"):
            if any(r[1].lower() == column.lower() for r in con.execute(f'PRAGMA table_info("{t}")')):
                col = next(r[1] for r in con.execute(f'PRAGMA table_info("{t}")') if r[1].lower() == column.lower())
                out += [(t, str(v)) for (v,) in con.execute(f'SELECT DISTINCT "{col}" FROM "{t}" WHERE "{col}" IS NOT '
                                                             f'NULL LIMIT 20000')]
    finally:
        con.close()
    return tuple(out)


def literal_issues(db_path: Path, sql: str) -> list[str]:
    """Messages for string literals compared with a column that does not store that value (LIKE patterns: no match)."""
    msgs = []
    for col, op, lit in _LIT.findall(sql or ""):
        name = col.split(".")[-1].strip('"`[]')
        vals = _stored_values(str(db_path), name)
        if not vals:
            continue
        stored = [v for _, v in vals]
        if op.upper() == "LIKE":
            pat = re.compile("^" + re.escape(lit).replace("%", ".*").replace("_", ".") + "$", re.I | re.S)
            if any(pat.match(v) for v in stored):
                continue
        elif lit in stored:
            continue
        low = lit.lower().strip("%")
        close = [v for v in stored if v.lower() == low] or [v for v in stored if low and low in v.lower()][:5] \
            or difflib.get_close_matches(lit, stored, n=5, cutoff=0.6)
        hint = ", ".join(repr(v) for v in close[:5]) if close else "none similar"
        msgs.append(f"'{lit}' is not a stored value of column {name}; closest stored values: {hint}.")
    return msgs


def gate_messages(db_path: Path, sql: str, res: bird.ExecResult) -> list[str]:
    if not res.ok:
        return [f"The query fails: {res.error}"]
    msgs = literal_issues(db_path, sql)
    if not res.rows:
        msgs.append("The query returns no rows. Check the filters and the stored values.")
    elif all(all(v is None for v in r) for r in res.rows):
        msgs.append("The query returns only NULL values. Check the joins and the columns.")
    return msgs


# ----------------------------------------------------------------------------------------------- answer
def messages(q: bird.Question, schema_text: str, cfg: Config2, history: list[tuple[str, str]] | None = None) -> list[dict]:
    system = SYSTEM + ("\n\n" + RULES if cfg.rules else "")
    user = f"Schema:\n{schema_text}\n\nEvidence: {q.evidence or '(none)'}\nQuestion: {q.question}"
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    for sql, fb in history or []:
        msgs += [{"role": "assistant", "content": f"```sql\n{sql}\n```"},
                 {"role": "user", "content": f"Checks on this query:\n{fb}\nWrite a corrected query."}]
    return msgs


def answer2(q: bird.Question, db_path: Path, schema_text: str, cfg: Config2, chat, gli=None) -> dict:
    t0 = time.time()
    log: dict[str, Any] = {"qid": q.qid, "config": cfg.name, "llm_calls": 0, "revisions": 0, "doubt_revisions": 0,
                           "schema_chars": len(schema_text)}
    base = messages(q, schema_text, cfg)
    smps = chat.samples(base, n=1, temperature=0.0)
    log["llm_calls"] += 1
    if cfg.n > 1:
        smps += chat.samples(base, n=cfg.n - 1, temperature=0.8)
        log["llm_calls"] += 1
    cands = []
    for k, s in enumerate(smps):
        sql = extract_sql(s.text)
        res = bird.execute(db_path, sql)
        msgs = gate_messages(db_path, sql, res) if cfg.gates else []
        hist: list[tuple[str, str]] = []
        for _ in range(cfg.max_rev if cfg.gates else 0):
            if not msgs:
                break
            hist.append((sql, "\n".join(msgs)))
            sql = extract_sql(chat.samples(messages(q, schema_text, cfg, hist), n=1, temperature=0.0)[0].text)
            log["llm_calls"] += 1
            log["revisions"] += 1
            res = bird.execute(db_path, sql)
            msgs = gate_messages(db_path, sql, res)
        if cfg.doubt and gli is not None and not msgs and k == 0:
            text = f"Question: {q.question}\nEvidence: {q.evidence or '(none)'}\nSQL: {sql}\nResult: {bird.preview(res)}"
            if float(gli.scores([text], [DOUBT_LABEL])[0, 0]) < cfg.doubt_threshold:
                fb = ("An automatic checker doubts that this result answers the question. Re-read the question: are "
                      "the returned columns exactly what is asked, and is every condition applied? If the query is "
                      "already right, return it unchanged.")
                sql2 = extract_sql(chat.samples(messages(q, schema_text, cfg, hist + [(sql, fb)]), n=1,
                                                temperature=0.0)[0].text)
                log["llm_calls"] += 1
                res2 = bird.execute(db_path, sql2)
                if not gate_messages(db_path, sql2, res2):
                    sql, res = sql2, res2
                    log["doubt_revisions"] += 1
        cands.append({"k": k, "sql": sql, "res": res, "pass": not msgs})
    pool = [c for c in cands if c["pass"] and c["res"].ok] or [c for c in cands if c["res"].ok] or cands
    votes = Counter(c["res"].key for c in pool)
    pick = max(pool, key=lambda c: (votes[c["res"].key], c["k"] == 0))
    log.update(sql=pick["sql"], exec_ok=pick["res"].ok, result_key=pick["res"].key, passed_gates=pick["pass"],
               seconds=round(time.time() - t0, 2))
    return log
