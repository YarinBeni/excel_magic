"""Q3: the LLM proposes hypotheses, the harness tests them on top of the frozen relational model.

The frozen FM (Kumo Relational graph layer + linear probe) gives a score per row. The LLM never submits scores. It
explores the database and writes hypotheses as SQL features ("drivers who failed to finish recently fail again"). For
each hypothesis the harness, not the LLM:
  1. computes the feature for the probe rows (training period), the validation rows and the test rows (table `rows`);
  2. checks for leakage deterministically: the feature of a validation row must not change when every record after a
     cut time is removed (search_path switched to filtered views);
  3. tests it: a logistic model on [FM logit + accepted features + new features] is fitted on the probe rows and scored
     on the later validation rows; the hypothesis is accepted only if the paired bootstrap gain in AUROC over the model
     without it is > 2 standard errors (and > 0.002);
  4. reports the result back. The workspace table `fm_train_errors` (probe rows only, never validation rows) shows
     where the FM is wrong, so hypotheses target what the FM misses.
DuckDB errors get a deterministic dialect hint. Final scores: the accepted model refitted on probe + validation rows;
with no accepted hypothesis the system returns the FM score unchanged.
"""
from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from typing import Any

import numpy as np
import pandas as pd

HINTS = [(r"julianday|datediff|timestampdiff", "DuckDB: days between two times = date_diff('day', a, b)."),
         (r"dateadd|date_add|date_sub", "DuckDB: shift a time with t - INTERVAL 30 DAY."),
         (r"INTERVAL.*(INTEGER|FLOAT|DOUBLE|BIGINT)|Cannot compare values of type INTERVAL",
          "DuckDB: turn an interval into days with date_diff('day', a, b) or epoch(x)/86400."),
         (r"must appear in the GROUP BY", "Add the column to GROUP BY or wrap it in any_value(...)."),
         (r"non-inner join on subquery|correlated", "Rewrite the correlated subquery as a LEFT JOIN on the entity with "
                                                   "the time condition in the ON clause, then GROUP BY r.<entity>, r.<time>."),
         (r"aggregate function calls cannot (contain window|be nested)", "Compute the inner aggregate in a CTE first."),
         (r"Ambiguous reference", "Qualify every column with its table alias.")]

SYSTEM = """You are a data scientist improving a prediction with hypotheses. A frozen relational model already scores
every row; your job is to find signal it misses.
Question: {question}
Tables (name: columns): {tables}
Workspace: `fm_train_errors` ({ent}, {time}, {target}, fm) has training-period rows with the label and the model's
score: look at rows where the model is wrong (high fm and label 0, low fm and label 1) and ask what they have in common.
Tools: sql(query) to explore; test_hypothesis(hypothesis, feature_sql) to test one idea; finish() when done.
feature_sql must SELECT r.{ent}, r.{time}, <one or more numeric feature columns> FROM rows r ... with exactly one row
per row of `rows` (use LEFT JOIN ... GROUP BY r.{ent}, r.{time}). Use only records strictly before r.{time}.
The harness computes the feature for training, validation and test rows, checks it for leakage, and keeps it only if it
improves on the model's score on later data. Test one clear idea at a time; {budget} tests at most."""

TOOLS = [{"type": "function", "function": {"name": "sql", "description": "Run DuckDB SQL (shows up to 15 rows).",
                                           "parameters": {"type": "object", "properties": {"query": {"type": "string"}},
                                                          "required": ["query"]}}},
         {"type": "function", "function": {"name": "test_hypothesis",
                                           "description": "Test one hypothesis given as a feature SQL over table rows.",
                                           "parameters": {"type": "object", "properties": {
                                               "hypothesis": {"type": "string"}, "feature_sql": {"type": "string"}},
                                               "required": ["hypothesis", "feature_sql"]}}},
         {"type": "function", "function": {"name": "finish", "description": "Stop testing.",
                                           "parameters": {"type": "object", "properties": {}}}}]


def hint(err: str) -> str:
    return " ".join(h for pat, h in HINTS if re.search(pat, err, re.I))


def _logit(p):
    p = np.clip(np.asarray(p, float), 1e-6, 1 - 1e-6)
    return np.log(p / (1 - p))


def _auc(y, s):
    from sklearn.metrics import roc_auc_score
    try:
        return float(roc_auc_score(y, s))
    except ValueError:
        return float("nan")


@dataclass
class Stacker:
    """Logistic model on [FM logit, quantile-normalised features, missing flags]; fitted on one frame, scores another."""
    C: float = 0.3

    def fit_predict(self, Ftr: pd.DataFrame, ytr, Fev: pd.DataFrame) -> np.ndarray:
        from sklearn.linear_model import LogisticRegression
        from sklearn.preprocessing import QuantileTransformer

        Xtr, Xev = [_logit(Ftr["fm"])[:, None]], [_logit(Fev["fm"])[:, None]]
        for c in [c for c in Ftr.columns if c != "fm"]:
            a, b = pd.to_numeric(Ftr[c], errors="coerce").to_numpy(float), pd.to_numeric(Fev[c], errors="coerce").to_numpy(float)
            fill = np.nanmedian(a) if np.isfinite(a).any() else 0.0
            qt = QuantileTransformer(n_quantiles=min(200, max(10, len(a))), output_distribution="normal")
            qt.fit(np.where(np.isfinite(a), a, fill)[:, None])
            Xtr += [qt.transform(np.where(np.isfinite(a), a, fill)[:, None]), (~np.isfinite(a)).astype(float)[:, None]]
            Xev += [qt.transform(np.where(np.isfinite(b), b, fill)[:, None]), (~np.isfinite(b)).astype(float)[:, None]]
        Xtr, Xev = np.hstack(Xtr), np.hstack(Xev)
        if len(np.unique(ytr)) < 2:
            return Fev["fm"].to_numpy(float)
        return LogisticRegression(C=self.C, max_iter=2000).fit(Xtr, ytr).predict_proba(Xev)[:, 1]


@dataclass
class HypoLoop:
    env: Any                      # RelEnv after setup() and relational_predict()
    chat: Any                     # vfd.llm.Chat
    budget: int = 10
    max_turns: int = 30
    min_gain: float = 0.002
    z: float = 2.0
    n_boot: int = 300
    accepted: list = field(default_factory=list)   # (name, frames dict)
    history: list = field(default_factory=list)

    # ------------------------------------------------------------------------------------------- feature plumbing
    def _run_feature(self, sql: str, rows: pd.DataFrame, con=None) -> pd.DataFrame:
        e = self.env
        if con is None:
            e.con.register("rows", rows[[e.ent, e.tcol]])
            f = e._q(sql, 300)
        else:
            con.register("rows", rows[[e.ent, e.tcol]])
            f = con.execute(sql).fetchdf()
        if e.ent not in f.columns or e.tcol not in f.columns:
            raise ValueError(f"feature_sql must return columns {e.ent}, {e.tcol} and the features")
        feats = [c for c in f.columns if c not in (e.ent, e.tcol)]
        if not feats:
            raise ValueError("feature_sql returned no feature column")
        f = f.drop_duplicates([e.ent, e.tcol])
        f[e.tcol] = pd.to_datetime(f[e.tcol])
        r = rows[[e.ent, e.tcol]].copy()
        r[e.tcol] = pd.to_datetime(r[e.tcol])
        out = r.merge(f, on=[e.ent, e.tcol], how="left")[feats]
        return out.apply(pd.to_numeric, errors="coerce").reset_index(drop=True)

    def _leaks(self, sql: str) -> float:
        """Share of validation rows (before the cut) whose feature changes when records after the cut are removed."""
        import duckdb

        e = self.env
        val = e._fm["val"]
        t = pd.to_datetime(val[e.tcol])
        cut = t.quantile(0.5)
        sub = val[t <= cut]
        if len(sub) < 20:
            return 0.0
        full = self._run_feature(sql, sub)
        if getattr(self, "_cut_con", None) is None:
            self._cut_con = duckdb.connect()
            for name, tab in e.db.table_dict.items():
                df = tab.df
                if tab.time_col:
                    df = df[pd.to_datetime(df[tab.time_col]) <= cut]
                self._cut_con.register(name, df)
            tl = e.train[[e.ent, e.tcol, e.tgt]]
            self._cut_con.register("train_labels", tl[pd.to_datetime(tl[e.tcol]) <= cut])
        part = self._run_feature(sql, sub, self._cut_con)
        a, b = full.to_numpy(float), part.to_numpy(float)
        same = np.isclose(a, b, equal_nan=True, rtol=1e-6, atol=1e-9).all(1)
        return float(1 - same.mean())

    def _frames(self, extra: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        fm = self.env._fm
        parts = {k: [fm[k][["fm"]]] for k in ("prb", "val", "test")}
        for i, fr in enumerate(extra):
            for k in parts:
                parts[k].append(fr[k].add_prefix(f"h{i}_"))
        return tuple(pd.concat(parts[k], axis=1) for k in ("prb", "val", "test"))

    def _gain(self, frames: list[dict]) -> tuple[float, float, float]:
        e = self.env
        y_prb, y_val = e._fm["prb"][e.tgt].astype(int).to_numpy(), e._fm["val"][e.tgt].astype(int).to_numpy()
        Pa, Va, _ = self._frames([f for _, f in self.accepted])
        Pb, Vb, _ = self._frames([f for _, f in self.accepted] + frames)
        sa = Stacker().fit_predict(Pa, y_prb, Va)
        sb = Stacker().fit_predict(Pb, y_prb, Vb)
        rng = np.random.default_rng(0)
        d = []
        for _ in range(self.n_boot):
            i = rng.integers(0, len(y_val), len(y_val))
            if len(np.unique(y_val[i])) == 2:
                d.append(_auc(y_val[i], sb[i]) - _auc(y_val[i], sa[i]))
        return _auc(y_val, sb) - _auc(y_val, sa), float(np.std(d)) if d else float("inf"), _auc(y_val, sb)

    # ------------------------------------------------------------------------------------------- tools
    def test_hypothesis(self, hypothesis: str, sql: str) -> str:
        e = self.env
        if len(self.history) >= self.budget:
            return "Budget used up: call finish()."
        try:
            fr = {k: self._run_feature(sql, e._fm[k]) for k in ("prb", "val", "test")}
            leak = self._leaks(sql)
        except Exception as ex:
            msg = f"ERROR {type(ex).__name__}: {str(ex)[:300]}"
            return msg + (" Hint: " + hint(msg) if hint(msg) else "")
        rec = {"hypothesis": hypothesis[:300], "sql": sql[:2000], "n_features": fr["prb"].shape[1],
               "missing": round(float(fr["prb"].isna().mean().mean()), 3), "leak": round(leak, 3)}
        if leak > 0.01:
            rec.update(decision="rejected: leakage")
            self.history.append(rec)
            return (f"REJECTED (leakage): the feature changes for {leak:.0%} of validation rows when records after the "
                    f"row time are removed. Filter every joined table on its time column < r.{e.tcol}.")
        if fr["prb"].nunique().max() <= 1:
            rec.update(decision="rejected: constant")
            self.history.append(rec)
            return "REJECTED: the feature is constant on the training rows."
        g, se, auc = self._gain([fr])
        ok = g > max(self.min_gain, self.z * se)
        rec.update(gain=round(g, 4), se=round(se, 4), val_auroc=round(auc, 4), decision="accepted" if ok else "rejected")
        self.history.append(rec)
        if ok:
            self.accepted.append((hypothesis[:80], fr))
        return (f"{'ACCEPTED' if ok else 'REJECTED'}: validation AUROC gain {g:+.4f} (se {se:.4f}) over the model "
                f"{'plus ' + str(len(self.accepted) - ok) + ' accepted hypotheses' if self.accepted else 'score alone'}"
                f"; missing {rec['missing']:.0%}. Tests left: {self.budget - len(self.history)}.")

    def run(self, question: str) -> dict:
        e = self.env
        e.con.register("fm_train_errors", e._fm["prb"])
        msgs = [{"role": "system", "content": SYSTEM.format(
                    question=question, ent=e.ent, time=e.tcol, target=e.tgt, budget=self.budget,
                    tables="; ".join(f"{n}: {', '.join(c[:25])}" for n, c in e.tables.items())[:6000])},
                {"role": "user", "content": "Begin: look at the workspace, then test your best hypotheses."}]
        t0, calls = time.time(), 0
        for _ in range(self.max_turns):
            r = self.chat.client.chat.completions.create(model=self.chat.model, messages=msgs, tools=TOOLS,
                                                         temperature=0.3, max_tokens=2000)
            calls += 1
            m = r.choices[0].message
            msgs.append({"role": "assistant", "content": m.content or "",
                         **({"tool_calls": [tc.model_dump() for tc in m.tool_calls]} if m.tool_calls else {})})
            if not m.tool_calls:
                msgs.append({"role": "user", "content": "Use a tool: sql, test_hypothesis or finish."})
                continue
            stop = False
            for tc in m.tool_calls:
                try:
                    a = json.loads(tc.function.arguments or "{}")
                except json.JSONDecodeError:
                    a = {}
                if tc.function.name == "sql":
                    try:
                        out = e.sql(a.get("query", ""))
                    except Exception as ex:
                        out = f"ERROR {type(ex).__name__}: {str(ex)[:300]}"
                        out += (" Hint: " + hint(out)) if hint(out) else ""
                elif tc.function.name == "test_hypothesis":
                    out = self.test_hypothesis(a.get("hypothesis", ""), a.get("feature_sql", ""))
                else:
                    out, stop = "finished", True
                msgs.append({"role": "tool", "tool_call_id": tc.id, "content": str(out)[:4000]})
            if stop or len(self.history) >= self.budget:
                break
        return self.finalize(calls, time.time() - t0)

    def finalize(self, calls: int, seconds: float) -> dict:
        e = self.env
        fm = e._fm
        P, V, T = self._frames([f for _, f in self.accepted])
        if self.accepted:
            PV = pd.concat([P, V], ignore_index=True)
            y = np.concatenate([fm["prb"][e.tgt].astype(int).to_numpy(), fm["val"][e.tgt].astype(int).to_numpy()])
            score = Stacker().fit_predict(PV, y, T)
        else:
            score = fm["test"]["fm"].to_numpy(float)
        auc_sys = float(e.task.evaluate(score).get("roc_auc", np.nan))
        auc_fm = float(e.task.evaluate(fm["test"]["fm"].to_numpy(float)).get("roc_auc", np.nan))
        return {"auroc": auc_sys, "auroc_fm": auc_fm, "accepted": [n for n, _ in self.accepted],
                "tested": len(self.history), "history": self.history, "llm_calls": calls, "seconds": round(seconds, 1)}
