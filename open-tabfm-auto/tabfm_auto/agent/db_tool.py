"""``tabfm-db``: read-only database exploration + table materialisation tool for the DB agent.

  tabfm-db schema      --db shop.sqlite                      # tables, columns, row counts, FK guesses
  tabfm-db sql         --db shop.sqlite --query "..." [--limit 20]
  tabfm-db materialize --db shop.sqlite --sql-file build_table.sql --target churned --out train.parquet
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, inspect, text

HIDDEN_PREFIX = "_"  # tables starting with "_" are harness-only (ground truth, meta) and invisible to the agent
FORBIDDEN = re.compile(r"\b(insert|update|delete|drop|alter|create|attach|pragma|replace|vacuum)\b", re.I)


def engine_for(db: str):
    url = db if "://" in db else f"sqlite:///{Path(db).resolve()}"
    return create_engine(url)


def _visible_tables(eng) -> list[str]:
    return [t for t in inspect(eng).get_table_names() if not t.startswith(HIDDEN_PREFIX)]


def _check_query(q: str) -> None:
    if FORBIDDEN.search(q):
        raise SystemExit("only read-only SELECT / WITH queries are allowed")
    if re.search(r"\b_[a-z]", q, re.I) and re.search(r"\bfrom\s+_|\bjoin\s+_", q, re.I):
        raise SystemExit("tables starting with '_' are not accessible")


def cmd_schema(a) -> None:
    eng = engine_for(a.db)
    insp = inspect(eng)
    out = {}
    with eng.connect() as con:
        for t in _visible_tables(eng):
            cols = insp.get_columns(t)
            n = con.execute(text(f'select count(*) from "{t}"')).scalar()
            sample = pd.read_sql(text(f'select * from "{t}" limit 3'), con)
            fks = [c["name"] for c in cols if c["name"].endswith("_id") and c["name"] != f"{t[:-1]}_id"]
            out[t] = {"rows": n, "columns": {c["name"]: str(c["type"]) for c in cols},
                      "likely_foreign_keys": fks, "sample": sample.to_dict("records")}
    print(json.dumps(out, indent=2, default=str))


def cmd_sql(a) -> None:
    _check_query(a.query)
    eng = engine_for(a.db)
    with eng.connect() as con:
        df = pd.read_sql(text(a.query), con)
    if a.limit:
        df = df.head(a.limit)
    with pd.option_context("display.max_columns", 50, "display.width", 200):
        print(df.to_string(index=False))
    print(f"[{len(df)} rows shown]")


def cmd_materialize(a) -> None:
    q = Path(a.sql_file).read_text()
    _check_query(q)
    eng = engine_for(a.db)
    with eng.connect() as con:
        df = pd.read_sql(text(q), con)
    if a.target not in df.columns:
        raise SystemExit(f"target column {a.target!r} not in query result columns {list(df.columns)}")
    if df[a.target].isna().any():
        raise SystemExit("target column has NULLs")
    if a.id_col and a.id_col in df.columns and df[a.id_col].duplicated().any():
        raise SystemExit(f"{a.id_col} is not unique: one row per entity is required")
    df.to_parquet(a.out, index=False)
    summary = {"rows": len(df), "columns": list(df.columns), "target": a.target,
               "target_summary": df[a.target].describe().to_dict(),
               "dtypes": {c: str(t) for c, t in df.dtypes.items()}}
    print(json.dumps(summary, indent=2, default=str))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="tabfm-db", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("schema"); s.add_argument("--db", required=True); s.set_defaults(fn=cmd_schema)
    s = sub.add_parser("sql"); s.add_argument("--db", required=True); s.add_argument("--query", required=True)
    s.add_argument("--limit", type=int, default=30); s.set_defaults(fn=cmd_sql)
    s = sub.add_parser("materialize"); s.add_argument("--db", required=True); s.add_argument("--sql-file", required=True)
    s.add_argument("--target", required=True); s.add_argument("--out", default="train.parquet")
    s.add_argument("--id-col", default=None); s.set_defaults(fn=cmd_materialize)
    a = ap.parse_args(argv)
    a.fn(a)
    return 0


if __name__ == "__main__":
    sys.exit(main())
