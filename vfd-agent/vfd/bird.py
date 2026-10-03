"""BIRD text-to-SQL data: questions (Arcwise-Plat, the fully corrected 498-question mini-dev set), SQLite databases,
schema rendering with column descriptions, safe execution and BIRD's result-set equality."""
from __future__ import annotations

import csv
import hashlib
import json
import os
import sqlite3
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

REVISQL_COMMIT = "9fac371aa22019e9912dcbd6572e8fe8194d352a"
ARCWISE_URL = f"https://raw.githubusercontent.com/uiuc-kang-lab/ReViSQL/{REVISQL_COMMIT}/data/arcwise_plat_full.json"


@dataclass
class Question:
    qid: int
    db_id: str
    question: str
    evidence: str
    gold_sql: str


@dataclass
class Column:
    table: str
    name: str
    type: str
    description: str = ""
    values: str = ""


@dataclass
class Schema:
    db_id: str
    tables: dict[str, list[Column]] = field(default_factory=dict)
    fks: list[tuple[str, str, str, str]] = field(default_factory=list)

    @property
    def columns(self) -> list[Column]:
        return [c for cols in self.tables.values() for c in cols]


def load_questions(path: str | None = None, cache_dir: str = "data") -> list[Question]:
    """Arcwise-Plat-Full (ReViSQL, pinned commit). Downloaded once into cache_dir; not redistributed with this repo."""
    p = Path(path) if path else Path(cache_dir) / "arcwise_plat_full.json"
    if not p.exists():
        import requests

        p.parent.mkdir(parents=True, exist_ok=True)
        r = requests.get(ARCWISE_URL, timeout=60)
        r.raise_for_status()
        p.write_bytes(r.content)
    return [Question(int(d["question_id"]), d["db_id"], d["question"], d.get("evidence") or "", d["SQL"])
            for d in json.loads(p.read_text())]


def find_db(root: str | os.PathLike, db_id: str) -> Path:
    """Locate <db_id>.sqlite under root (the mini-dev zip nests it as .../dev_databases/<db_id>/<db_id>.sqlite)."""
    hits = sorted(Path(root).rglob(f"{db_id}.sqlite"))
    if not hits:
        raise FileNotFoundError(f"{db_id}.sqlite not found under {root}")
    return hits[0]


def _descriptions(db_dir: Path, table: str) -> dict[str, tuple[str, str]]:
    """BIRD database_description/<table>.csv: original_column_name, column_name, column_description, ..., value_description."""
    for cand in (db_dir / "database_description" / f"{table}.csv", db_dir / "database_description" / f"{table.lower()}.csv"):
        if cand.exists():
            out = {}
            raw = cand.read_bytes()
            for enc in ("utf-8-sig", "latin-1"):
                try:
                    rows = list(csv.DictReader(raw.decode(enc).splitlines()))
                    break
                except UnicodeDecodeError:
                    continue
            for r in rows:
                name = (r.get("original_column_name") or "").strip()
                if name:
                    desc = (r.get("column_description") or r.get("column_name") or "").strip()
                    out[name.lower()] = (desc, (r.get("value_description") or "").strip())
            return out
    return {}


def load_schema(db_path: Path, n_values: int = 3) -> Schema:
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode(errors="replace")
    s = Schema(db_path.stem)
    tables = [r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
    for t in tables:
        desc = _descriptions(db_path.parent, t)
        cols = []
        for _, name, typ, *_ in con.execute(f'PRAGMA table_info("{t}")'):
            try:
                vals = [str(v[0])[:40] for v in con.execute(f'SELECT DISTINCT "{name}" FROM "{t}" WHERE "{name}" IS NOT NULL LIMIT {n_values}')]
            except sqlite3.Error:
                vals = []
            d, vd = desc.get(name.lower(), ("", ""))
            cols.append(Column(t, name, typ or "", d, ", ".join(vals) + (f" ({vd[:120]})" if vd else "")))
        s.tables[t] = cols
        for row in con.execute(f'PRAGMA foreign_key_list("{t}")'):
            s.fks.append((t, row[3], row[2], row[4]))
    con.close()
    return s


def render_schema(s: Schema, columns: set[tuple[str, str]] | None = None) -> str:
    """CREATE TABLE text with a comment per column (description and example values). `columns` restricts to a subset
    (schema linking); a table is kept if any of its columns is kept."""
    out = []
    for t, cols in s.tables.items():
        keep = [c for c in cols if columns is None or (t.lower(), c.name.lower()) in columns]
        if not keep:
            continue
        lines = []
        for c in keep:
            note = "; ".join(x for x in (c.description, f"e.g. {c.values}" if c.values else "") if x)
            lines.append(f'  "{c.name}" {c.type}' + (f"  -- {note}" if note else ""))
        out.append(f'CREATE TABLE "{t}" (\n' + ",\n".join(lines) + "\n);")
    fks = [f"-- {a}.{b} = {c}.{d}" for a, b, c, d in s.fks
           if columns is None or ((a.lower(), b.lower()) in columns and (c.lower(), (d or "").lower()) in columns)]
    return "\n".join(out + fks)


@dataclass
class ExecResult:
    ok: bool
    rows: list[tuple] | None
    error: str = ""
    seconds: float = 0.0

    @property
    def key(self) -> str:
        """Order-insensitive hash of the result set (BIRD compares set(rows))."""
        if not self.ok:
            return "ERR"
        return hashlib.sha1(repr(sorted({tuple(map(_norm, r)) for r in self.rows}, key=repr)).encode()).hexdigest()[:16]


def _norm(v: Any) -> Any:
    if isinstance(v, float):
        return round(v, 6)
    return v


def execute(db_path: Path, sql: str, timeout_s: float = 30.0, max_rows: int = 100_000) -> ExecResult:
    """Read-only execution with a wall-clock limit (sqlite progress handler aborts long queries)."""
    t0 = time.time()
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    con.text_factory = lambda b: b.decode(errors="replace")
    con.set_progress_handler(lambda: 1 if time.time() - t0 > timeout_s else 0, 10_000)
    try:
        cur = con.execute(sql)
        rows = cur.fetchmany(max_rows)
        return ExecResult(True, [tuple(r) for r in rows], seconds=time.time() - t0)
    except Exception as e:  # sqlite3.Error, OperationalError on interrupt, Warning for multiple statements
        return ExecResult(False, None, f"{type(e).__name__}: {str(e)[:200]}", time.time() - t0)
    finally:
        con.close()


def same_result(a: ExecResult, b: ExecResult) -> bool:
    return a.ok and b.ok and a.key == b.key


def preview(r: ExecResult, n: int = 5, width: int = 200) -> str:
    if not r.ok:
        return f"ERROR {r.error}"
    head = "; ".join(", ".join(str(v)[:30] for v in row) for row in r.rows[:n])
    return f"{len(r.rows)} rows: {head}"[:width]
