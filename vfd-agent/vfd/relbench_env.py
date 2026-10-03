"""Claims C4 / C7: prediction questions over a whole relational database, answered by an LLM agent.

The agent sees a RelBench database in DuckDB (every table, cut at the test timestamp), a table `train_labels`
(entity, time, label) and a table `test_rows` (entity, time), and a question in plain words ("which drivers will not
finish a race in the next month?"). It must submit a score for every test row. The official RelBench evaluator grades
the submission; the agent never sees test labels.

Tools by config:
  D   sql, submit                                   the LLM alone: heuristics written in SQL
  E   + deep_fit_predict(train_sql, test_sql)       Kumo Tabular-L fitted in context on a feature table the LLM builds
  R   + relational_predict()                         Kumo Relational graph layer + linear probe (paper 2), no SQL needed
  ER  both
"""
from __future__ import annotations

import json
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

QUESTIONS = {
    ("rel-f1", "driver-dnf"): "For each driver in test_rows, how likely is it that the driver does not finish (DNF) a race in the next month after `date`?",
    ("rel-f1", "driver-top3"): "For each driver in test_rows, how likely is a top-3 qualifying finish in the next month after `date`?",
    ("rel-trial", "study-outcome"): "For each clinical trial in test_rows, how likely is it that the trial achieves its primary outcome?",
    ("rel-event", "user-repeat"): "For each user in test_rows, how likely is the user to attend another event in the next 7 days after `timestamp`?",
    ("rel-event", "user-ignore"): "For each user in test_rows, how likely is the user to ignore more than 2 event invitations in the next 7 days?",
    ("rel-avito", "user-visits"): "For each user in test_rows, how likely is the user to visit more than one ad in the next 4 days?",
    ("rel-avito", "user-clicks"): "For each user in test_rows, how likely is the user to click more than one ad in the next 4 days?",
    ("rel-hm", "user-churn"): "For each customer in test_rows, how likely is it that the customer makes no transaction in the next week (churn)?",
}

SYSTEM = """You are a data scientist answering a prediction question over a relational database in DuckDB.
Question: {question}
Tables (name: columns): {tables}
`train_labels` has past examples ({ent}, {time}, {target}) with the known answer; `test_rows` has ({ent}, {time}) to score.
Use only data with time <= the row's {time} when you build features (no peeking into the future).
Finish by calling submit(sql) with a query that returns columns {ent}, {time}, score for EVERY row of test_rows
(higher score = more likely). {extra}"""
EXTRA = {"E": "deep_fit_predict fits a frozen tabular foundation model on a feature table you define with SQL "
              "(one row per train_labels row, with the label column) and scores a matching feature table for test_rows.",
         "R": "relational_predict scores test_rows with a frozen relational foundation model that reads the whole "
              "database graph around each entity; it needs no SQL and returns a validation AUROC.",
         }


@dataclass
class RelEnv:
    dataset: str
    task_name: str
    max_test: int = 0
    seed: int = 0
    device: str = "cuda"
    log: list = field(default_factory=list)

    def setup(self):
        import duckdb
        import relbench

        self.task = relbench.load_dataset(self.dataset).load_task(self.task_name)
        self.db = self.task.get_db(upto_test_timestamp=False)
        t = self.task
        self.ent, self.tcol, self.tgt = t.entity_col, t.time_col, t.target_col
        self.train = t.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
        self.val = t.get_table("val", mask_input_cols=False).df.reset_index(drop=True)
        self.test = t.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
        cutoff = pd.Timestamp(self.test[self.tcol].max())
        self.con = duckdb.connect()
        self.tables = {}
        for name, tab in self.db.table_dict.items():
            df = tab.df
            if tab.time_col:
                df = df[pd.to_datetime(df[tab.time_col]) <= cutoff]
            self.con.register(name, df)
            self.tables[name] = list(df.columns)
        self.con.register("train_labels", self.train[[self.ent, self.tcol, self.tgt]])
        self.con.register("test_rows", self.test[[self.ent, self.tcol]])
        self.submission = None
        return self

    def system_prompt(self, cfg: str) -> str:
        extra = " ".join(EXTRA[k] for k in ("E", "R") if k in cfg)
        tabs = "; ".join(f"{n}: {', '.join(c[:25])}" for n, c in self.tables.items())
        return SYSTEM.format(question=QUESTIONS[(self.dataset, self.task_name)], tables=tabs[:6000], ent=self.ent,
                             time=self.tcol, target=self.tgt, extra=extra)

    # ----------------------------------------------------------------------------------------------- tools
    def sql(self, query: str, n: int = 15) -> str:
        res = self.con.execute(query).fetchdf()
        return f"{len(res)} rows; columns {list(res.columns)}\n{res.head(n).to_string(max_colwidth=50)}"

    def reset(self):
        """New episode on the same task: forget the submission and hide every prediction table."""
        for name in ("pred_deep", "pred_relational"):
            try:
                self.con.unregister(name)
            except Exception:
                pass
        self.submission = None
        self.log = []

    def _align(self, df: pd.DataFrame) -> tuple[np.ndarray, float]:
        df = df.copy()
        df[self.tcol] = pd.to_datetime(df[self.tcol])
        t = self.test[[self.ent, self.tcol]].copy()
        t[self.tcol] = pd.to_datetime(t[self.tcol])
        m = t.merge(df[[self.ent, self.tcol, "score"]].drop_duplicates([self.ent, self.tcol]), on=[self.ent, self.tcol], how="left")
        s = pd.to_numeric(m["score"], errors="coerce")
        return s.fillna(s.mean() if s.notna().any() else 0.0).to_numpy(float), float(s.notna().mean())

    def submit(self, query: str) -> str:
        df = self.con.execute(query).fetchdf()
        missing = {self.ent, self.tcol, "score"} - set(df.columns)
        if missing:
            return f"REJECTED: the query must return columns {self.ent}, {self.tcol}, score (missing {missing})"
        scores, cover = self._align(df)
        if cover < 0.5:
            return f"REJECTED: only {cover:.0%} of test_rows have a score; join on {self.ent} and {self.tcol}."
        self.submission = {"scores": scores, "coverage": cover}
        return f"submitted ({cover:.0%} of test rows scored)"

    def deep_fit_predict(self, train_sql: str, test_sql: str, spec: str) -> str:
        sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
        from .deep import DeepTool
        from .tabular import encode_frames

        tr = self.con.execute(train_sql).fetchdf()
        te = self.con.execute(test_sql).fetchdf()
        if self.tgt not in tr.columns:
            return f"train_sql must return the label column {self.tgt}"
        feats = [c for c in tr.columns if c not in (self.ent, self.tcol, self.tgt) and c in te.columns]
        if not feats:
            return "no shared feature columns between train_sql and test_sql"
        dt = DeepTool(model=spec, max_context=10_000)
        q = dt.predict(tr[[self.tcol] + feats + [self.tgt]], self.tgt, time_col=self.tcol)  # held-out quality
        rng = np.random.default_rng(self.seed)
        Xtr, Xte = encode_frames(tr[feats], te[feats])
        p = dt._fit_predict(spec, "binary", Xtr, tr[self.tgt].astype(int).to_numpy(), Xte, rng)
        out = te[[self.ent, self.tcol]].copy()
        out["score"] = p[:, 1]
        self.con.register("pred_deep", out)
        return (f"fitted on {len(tr)} rows x {len(feats)} features; time-split held-out {q['metric']} {q['value']}"
                + (f" (LightGBM {q['check_value']})" if "check_value" in q else "")
                + "; scores in table pred_deep(" + f"{self.ent}, {self.tcol}, score)")

    def relational_predict(self, n_ctx: int = 512, n_train: int = 4000) -> str:
        sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "frozen-embeddings-retrieval"))
        import torch
        from fer.relbench_layers import Subgrapher, embed_rows, probe_scores_multi, safe_auc, time_split
        from sdm.models import KumoRelational

        if getattr(self, "_rel", None) is not None:  # computed once per task, re-exposed per episode
            self.con.register("pred_relational", self._rel_df)
            return self._rel
        rng = np.random.default_rng(self.seed)
        tr, va, te, ent, tc, tgt = self.train, self.val, self.test, self.ent, self.tcol, self.tgt
        sub = Subgrapher(self.db, self.task.entity_table, k_children=50)
        pool_ctx, pool_prb, _ = time_split(tr, tc, n_train)
        trs = pool_ctx.sample(frac=1.0, random_state=self.seed).drop_duplicates(ent)
        pos, neg = trs[trs[tgt] == 1], trs[trs[tgt] != 1]
        nc_pos = min(len(pos), max(n_ctx // 2, int(n_ctx * len(pos) / max(len(trs), 1))))
        ctx = pd.concat([pos.head(nc_pos), neg.head(n_ctx - nc_pos)]).reset_index(drop=True)
        prb = pool_prb.iloc[rng.choice(len(pool_prb), min(n_train, len(pool_prb)), replace=False)].reset_index(drop=True)
        if len(va) > 5000:
            va = va.iloc[np.sort(rng.choice(len(va), 5000, replace=False))].reset_index(drop=True)
        torch.manual_seed(self.seed)
        model = KumoRelational(device=self.device)
        model.eval()
        y_ctx = ctx[tgt].to_numpy().astype(int)
        Etr, _ = embed_rows(model, sub, ctx, y_ctx, prb, ent, tc, self.device)
        Eva, _ = embed_rows(model, sub, ctx, y_ctx, va, ent, tc, self.device)
        Ete, _ = embed_rows(model, sub, ctx, y_ctx, te, ent, tc, self.device)
        pv, pt = probe_scores_multi(Etr["gnn"], prb[tgt].to_numpy().astype(int), [Eva["gnn"], Ete["gnn"]])
        out = te[[ent, tc]].copy()
        out["score"] = pt["linear"]
        self.con.register("pred_relational", out)
        self._rel_df = out
        self._rel = (f"relational model scored all test rows; validation AUROC {safe_auc(va[tgt], pv['linear']):.3f}; "
                     f"scores in table pred_relational({ent}, {tc}, score)")
        return self._rel

    def evaluate(self) -> dict[str, Any]:
        if self.submission is None:
            return {"auroc": float("nan"), "submitted": False}
        r = self.task.evaluate(self.submission["scores"])
        return {"auroc": float(r.get("roc_auc", np.nan)), "submitted": True, "coverage": self.submission["coverage"]}


def tool_specs(cfg: str) -> list[dict]:
    def fn(name, desc, props, req):
        return {"type": "function", "function": {"name": name, "description": desc,
                                                 "parameters": {"type": "object", "properties": props, "required": req}}}
    S = {"type": "string"}
    specs = [fn("sql", "Run DuckDB SQL over the database tables, train_labels and test_rows.", {"query": S}, ["query"]),
             fn("submit", "Submit scores: a SQL query returning entity, time, score for every test row.", {"query": S}, ["query"])]
    if "E" in cfg:
        specs.append(fn("deep_fit_predict", "Fit a frozen tabular foundation model: train_sql returns one row per "
                        "train_labels row (entity, time, features..., label); test_sql returns one row per test_rows row "
                        "(entity, time, same features). Writes table pred_deep.", {"train_sql": S, "test_sql": S},
                        ["train_sql", "test_sql"]))
    if "R" in cfg:
        specs.append(fn("relational_predict", "Score every test row with a frozen relational foundation model over the "
                        "database graph. Writes table pred_relational.", {}, []))
    return specs


def run_agent(env: RelEnv, chat, cfg: str, deep_spec: str, max_steps: int = 25) -> dict:
    msgs = [{"role": "system", "content": env.system_prompt(cfg)}, {"role": "user", "content": "Begin."}]
    t0 = time.time()
    calls = 0
    for step in range(max_steps):
        r = chat.client.chat.completions.create(model=chat.model, messages=msgs, tools=tool_specs(cfg), temperature=0.2,
                                                max_tokens=2000)
        calls += 1
        m = r.choices[0].message
        msgs.append({"role": "assistant", "content": m.content or "",
                     **({"tool_calls": [tc.model_dump() for tc in m.tool_calls]} if m.tool_calls else {})})
        if not m.tool_calls:
            msgs.append({"role": "user", "content": "Use a tool; finish with submit(query)."})
            continue
        for tc in m.tool_calls:
            try:
                a = json.loads(tc.function.arguments or "{}")
            except json.JSONDecodeError:
                a = {}
            try:
                if tc.function.name == "sql":
                    out = env.sql(a.get("query", ""))
                elif tc.function.name == "submit":
                    out = env.submit(a.get("query", ""))
                elif tc.function.name == "deep_fit_predict" and "E" in cfg:
                    out = env.deep_fit_predict(a.get("train_sql", ""), a.get("test_sql", ""), deep_spec)
                elif tc.function.name == "relational_predict" and "R" in cfg:
                    out = env.relational_predict()
                else:
                    out = f"unknown tool {tc.function.name}"
            except Exception as e:
                out = f"ERROR {type(e).__name__}: {str(e)[:300]}"
            env.log.append({"step": step, "tool": tc.function.name, "args": {k: str(v)[:300] for k, v in a.items()},
                            "out": str(out)[:300]})
            msgs.append({"role": "tool", "tool_call_id": tc.id, "content": str(out)[:4000]})
        if env.submission is not None and any(x["tool"] == "submit" and x["out"].startswith("submitted") for x in env.log[-3:]):
            break
    res = env.evaluate()
    res.update(config=cfg, llm_calls=calls, seconds=round(time.time() - t0, 1),
               tools=dict(pd.Series([x["tool"] for x in env.log]).value_counts()) if env.log else {})
    return res
