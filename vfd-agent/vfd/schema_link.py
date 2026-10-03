"""Schema linking: rank a database's columns by relevance to a question; score against the columns the corrected gold
SQL uses. Scorers: lexical overlap (baseline), embedding cosine, cross-encoder reranker, GLiClass (columns as labels),
GLiNER bi-encoder (columns as entity types; a column's score is its best span score)."""
from __future__ import annotations

import re
import time

import numpy as np

from .bird import Schema


def gold_columns(sql: str, schema: Schema) -> set[tuple[str, str]]:
    """(table, column) pairs used by the SQL, lower-cased. Aliases resolved; an unqualified column is assigned to every
    table in the query that has it (usually one)."""
    import sqlglot
    from sqlglot import exp

    try:
        tree = sqlglot.parse_one(sql, read="sqlite")
    except Exception:
        return set()
    by_table = {t.lower(): {c.name.lower() for c in cols} for t, cols in schema.tables.items()}
    alias = {}
    used_tables = set()
    for t in tree.find_all(exp.Table):
        name = t.name.lower()
        used_tables.add(name)
        alias[(t.alias_or_name or name).lower()] = name
    out = set()
    for c in tree.find_all(exp.Column):
        col = c.name.lower()
        if c.table:
            tab = alias.get(c.table.lower(), c.table.lower())
            if col in by_table.get(tab, set()):
                out.add((tab, col))
        else:
            for tab in used_tables:
                if col in by_table.get(tab, set()):
                    out.add((tab, col))
    return out


def column_label(c) -> str:
    natural = re.sub(r"[_\s]+", " ", c.name).strip()
    desc = c.description if c.description and c.description.lower() != natural.lower() else ""
    return f"{c.table}.{natural}" + (f": {desc}" if desc else "")


def _tokens(s: str) -> set[str]:
    return {w for w in re.findall(r"[a-z0-9]+", s.lower().replace("_", " ")) if len(w) > 1}


def lexical_scores(query: str, labels: list[str]) -> np.ndarray:
    q = _tokens(query)
    return np.array([len(q & _tokens(lab)) / (1 + len(_tokens(lab))) ** 0.5 for lab in labels], np.float32)


class EmbedScorer:
    def __init__(self, model_id: str = "BAAI/bge-small-en-v1.5", device: str = "cuda"):
        from sentence_transformers import SentenceTransformer

        self.m = SentenceTransformer(model_id, device=device)

    def scores(self, query: str, labels: list[str]) -> np.ndarray:
        e = self.m.encode([query] + labels, normalize_embeddings=True, batch_size=128)
        return (e[1:] @ e[0]).astype(np.float32)


class RerankScorer:
    def __init__(self, model_id: str = "BAAI/bge-reranker-v2-m3", device: str = "cuda"):
        from sentence_transformers import CrossEncoder

        self.m = CrossEncoder(model_id, device=device, max_length=512)

    def scores(self, query: str, labels: list[str]) -> np.ndarray:
        return np.asarray(self.m.predict([(query, lab) for lab in labels], batch_size=64), np.float32)


class GLiClassLinker:
    def __init__(self, scorer, chunk: int = 40):
        self.s, self.chunk = scorer, chunk  # uni-encoder: labels share the 512-token window with the text

    def scores(self, query: str, labels: list[str]) -> np.ndarray:
        out = []
        for i in range(0, len(labels), self.chunk):
            out.append(self.s.scores([query], labels[i:i + self.chunk])[0])
        return np.concatenate(out)


class GLiNERLinker:
    def __init__(self, model_id: str = "knowledgator/gliner-bi-base-v2.0", device: str = "cuda", threshold: float = 0.05):
        from gliner import GLiNER

        self.m = GLiNER.from_pretrained(model_id).to(device)
        self.threshold = threshold

    def scores(self, query: str, labels: list[str]) -> np.ndarray:
        ents = self.m.predict_entities(query, labels, threshold=self.threshold)
        best = {}
        for e in ents:
            best[e["label"]] = max(best.get(e["label"], 0.0), float(e["score"]))
        return np.array([best.get(lab, 0.0) for lab in labels], np.float32)


def evaluate_linking(items, scorers: dict, ks=(5, 10, 20)) -> tuple[list[dict], dict]:
    """items: (qid, query, schema, gold set). Returns per-question rows and per-scorer seconds per question."""
    rows, lat = [], {}
    for name, sc in scorers.items():
        t0 = time.time()
        for qid, query, schema, gold in items:
            cols = schema.columns
            labels = [column_label(c) for c in cols]
            keys = [(c.table.lower(), c.name.lower()) for c in cols]
            s = sc.scores(query, labels) if not callable(sc) else sc(query, labels)
            order = np.argsort(-s, kind="stable")
            rank = {keys[j]: r for r, j in enumerate(order)}
            gr = sorted(rank.get(g, len(keys)) for g in gold)
            row = {"scorer": name, "qid": qid, "n_columns": len(keys), "n_gold": len(gold)}
            for k in ks:
                row[f"recall@{k}"] = float(np.mean([r < k for r in gr])) if gr else np.nan
                row[f"all@{k}"] = float(all(r < k for r in gr)) if gr else np.nan
            rows.append(row)
        lat[name] = (time.time() - t0) / max(1, len(items))
    return rows, lat
