"""The analyst agent (claims C3 / C6): an LLM tool loop over one table in DuckDB, with the three-model tools.

Tools offered depend on the config:
  sql(query)                      always: DuckDB over table `data`; returns columns, row count and the first rows
  deep_*(...)                     config E, F: the DEEP tool (drivers, predict, what_if, anomalies, drift) on a SQL-defined frame
  text_labels(column, labels)     config F: GLiClass labels of a text column, added as new columns `<col>__<label>`
  record_insight(text, evidence_sql)
                                  every config records insights; in F each one is verified first: the evidence SQL is
                                  re-run, every number in the text must appear in its result (2% tolerance), and
                                  GLiClass must find the text entailed by the result; rejected insights come back with
                                  the reason. Accepted insights form the ledger the LLM sees each turn (state).
  finish(summary)
Configs: D = sql + record (no checks); E = D + deep tools; F = E + text labels + verified ledger.
"""
from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from typing import Any

import numpy as np
import pandas as pd

SYSTEM = """You are a senior data analyst. Goal: {goal}
Role: {role}. Dataset: {desc}
The data is in a DuckDB table named `data` with columns: {columns}.
Work step by step with the tools. Each time you find a concrete, data-backed insight, call record_insight with one
sentence and the SQL that shows it. Look for distributions, differences between groups, trends over time, outliers,
and what drives an outcome. Record 5 to {max_insights} insights, then call finish(summary).
{extra}"""

EXTRA_DEEP = ("You also have deep_* tools: frozen tabular foundation models that read ALL rows and return drivers "
              "(which columns predict an outcome), what-if curves, anomalies and drift between periods. Use them for "
              "'why' and 'what drives' questions instead of guessing from a few GROUP BYs.")
EXTRA_TEXT = ("text_labels turns a free-text column into label columns you can then GROUP BY or use as drivers.")
EXTRA_LEDGER = ("Insights are checked before they are accepted: every number you state must appear in the result of "
                "your evidence SQL. Accepted insights so far:\n{ledger}")

NUM = re.compile(r"-?\d+(?:[.,]\d+)?%?")


def number_supported(tok: str, vals: list[float]) -> bool:
    """A number in an insight is supported if some value in the evidence equals it up to the rounding the text shows
    (half a unit of its last digit, or 0.5% relative). Percentages may match a fraction (x100). Small bare integers
    (0, 1, 2) are list positions or counts of things in the sentence, not claims, and are skipped."""
    pct = tok.endswith("%")
    t = tok.rstrip("%").replace(",", ".") if tok.count(",") == 1 and "." not in tok and len(tok.split(",")[1]) != 3 else \
        tok.rstrip("%").replace(",", "")
    v = float(t)
    if not pct and "." not in t and abs(v) < 3:
        return True
    d = len(t.split(".")[1]) if "." in t else 0
    cands = vals + ([c * 100 for c in vals] if pct else [])
    if pct:  # a share the analyst computed from two evidence numbers (6 of 372 -> 1.6%)
        vs = [x for x in vals[:60] if x > 0]
        cands += [100 * a / b for a in vs for b in vs if a <= b]
    return any(abs(v - c) <= max(0.005 * abs(c), 0.5 * 10 ** (-d)) for c in cands)


@dataclass
class AnalystConfig:
    name: str
    deep: bool = False
    text: bool = False
    verify: bool = False
    max_steps: int = 30
    max_insights: int = 8
    profile: bool = False           # v2: deterministic data profile in the prompt
    skill: bool = False             # v2: short analysis checklist (curated skill card)
    coverage: bool = False          # v2: GLiClass labels each insight; the reply lists analysis kinds still missing


CONFIGS = {"D": AnalystConfig("D"), "E": AnalystConfig("E", deep=True),
           "F": AnalystConfig("F", deep=True, text=True, verify=True),
           "P": AnalystConfig("P", profile=True, skill=True),
           "PC": AnalystConfig("PC", profile=True, skill=True, coverage=True),
           "PCE": AnalystConfig("PCE", profile=True, skill=True, coverage=True, deep=True)}

SKILL = """Analysis checklist (cover several kinds; every insight states the actual numbers from your query):
1. Shares: how the goal's main measure splits across the main categories.
2. Trend: how it changes over time (month or quarter), and when it changed most.
3. Groups: which group differs most from the others (averages, rates, counts).
4. Top and bottom: the entities (people, items, places) with the highest and lowest values.
5. Outliers: unusual rows or periods, and how unusual they are.
6. Relations: two measures that move together, or a factor that goes with a worse outcome."""

KINDS = ["share or distribution across categories", "trend or change over time", "difference between groups",
         "top or bottom entities", "outlier or unusual case", "relation between two measures"]


def profile_frame(con) -> str:
    """Deterministic profile of table `data`: per column type, distinct count, NULL share, range or frequent values."""
    summ = con.execute("SUMMARIZE data").fetchdf()
    lines = []
    for _, r in summ.iterrows():
        c, typ = r["column_name"], str(r["column_type"])
        part = f"- {c} ({typ}): ~{int(r['approx_unique'])} distinct, {float(r['null_percentage']):.0f}% null"
        if int(r["approx_unique"]) <= 12 or typ in ("VARCHAR",):
            top = con.execute(f'SELECT "{c}" v, COUNT(*) n FROM data GROUP BY 1 ORDER BY 2 DESC LIMIT 5').fetchall()
            part += "; top: " + ", ".join(f"{str(v)[:25]} ({n})" for v, n in top)
        else:
            part += f"; range {str(r['min'])[:25]} to {str(r['max'])[:25]}"
        lines.append(part)
    n = con.execute("SELECT COUNT(*) FROM data").fetchone()[0]
    return f"Data profile ({n} rows):\n" + "\n".join(lines)


def tool_specs(cfg: AnalystConfig) -> list[dict]:
    def fn(name, desc, props, req):
        return {"type": "function", "function": {"name": name, "description": desc,
                                                 "parameters": {"type": "object", "properties": props, "required": req}}}

    S = {"type": "string"}
    specs = [fn("sql", "Run a DuckDB SQL query on table `data`.", {"query": S}, ["query"]),
             fn("record_insight", "Record one insight with the SQL that shows it.", {"text": S, "evidence_sql": S},
                ["text", "evidence_sql"]),
             fn("finish", "Finish with a short summary of the insights.", {"summary": S}, ["summary"])]
    if cfg.deep:
        specs += [
            fn("deep_drivers", "Which columns drive / predict `target` (permutation importance of a frozen tabular model "
               "on held-out rows).", {"target": S, "exclude": {"type": "array", "items": S}}, ["target"]),
            fn("deep_what_if", "How the predicted `target` changes when `column` takes different values.",
               {"target": S, "column": S}, ["target", "column"]),
            fn("deep_anomalies", "Rows whose `target` is most surprising given the other columns.",
               {"target": S, "id_column": S}, ["target"]),
            fn("deep_drift", "Did the data change between rows before and after `cutoff` in `time_column`? Which columns?",
               {"time_column": S, "cutoff": S}, ["time_column", "cutoff"]),
            fn("deep_predict", "Held-out predictability of `target` (is there real signal?).",
               {"target": S, "time_column": S}, ["target"])]
    if cfg.text:
        specs.append(fn("text_labels", "Label a free-text column with your label names (zero-shot classifier); adds "
                        "0/1 columns `<column>__<label>` to `data` and returns label counts.",
                        {"column": S, "labels": {"type": "array", "items": S}}, ["column", "labels"]))
    return specs


@dataclass
class Analyst:
    chat: Any                       # vfd.llm.Chat
    cfg: AnalystConfig
    deep: Any = None                # vfd.deep.DeepTool
    gli: Any = None                 # vfd.signals.GLiClassScorer
    log: list = field(default_factory=list)

    def run(self, df: pd.DataFrame, meta: dict) -> dict:
        import duckdb

        self.con = duckdb.connect()
        self.df = df.copy()
        self.con.register("data", self.df)
        self.ledger: list[dict] = []
        self.rejected: list[dict] = []
        t0 = time.time()
        extra = "\n".join(x for x, on in ((EXTRA_DEEP, self.cfg.deep), (EXTRA_TEXT, self.cfg.text), (SKILL, self.cfg.skill))
                           if on)
        if self.cfg.profile:
            extra += "\n" + profile_frame(self.con)
        self.covered: set[str] = set()
        sys_text = SYSTEM.format(goal=meta.get("goal", ""), role=meta.get("role", "analyst"),
                                 desc=str(meta.get("dataset_description", ""))[:1500],
                                 columns=", ".join(f"{c} ({t})" for c, t in zip(self.df.columns, self.df.dtypes.astype(str))),
                                 max_insights=self.cfg.max_insights, extra=extra)
        msgs = [{"role": "system", "content": sys_text}, {"role": "user", "content": "Start the analysis."}]
        summary, calls = "", 0
        for step in range(self.cfg.max_steps):
            if self.cfg.verify and step:
                led = "\n".join(f"- {x['text']}" for x in self.ledger) or "(none yet)"
                msgs[0]["content"] = sys_text + "\n" + EXTRA_LEDGER.format(ledger=led)
            r = self.chat.client.chat.completions.create(model=self.chat.model, messages=msgs, tools=tool_specs(self.cfg),
                                                         temperature=0.2, max_tokens=1500)
            calls += 1
            m = r.choices[0].message
            msgs.append({"role": "assistant", "content": m.content or "", **(
                {"tool_calls": [tc.model_dump() for tc in m.tool_calls]} if m.tool_calls else {})})
            if not m.tool_calls:
                msgs.append({"role": "user", "content": "Use a tool, or call finish(summary)."})
                continue
            done = False
            for tc in m.tool_calls:
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except json.JSONDecodeError:
                    args = {}
                out = self.call(tc.function.name, args)
                self.log.append({"step": step, "tool": tc.function.name, "args": args, "out": str(out)[:400]})
                msgs.append({"role": "tool", "tool_call_id": tc.id, "content": str(out)[:4000]})
                if tc.function.name == "finish":
                    summary, done = args.get("summary", ""), True
            if done or len(self.ledger) >= self.cfg.max_insights + 4:
                break
        return {"config": self.cfg.name, "insights": [x["text"] for x in self.ledger], "summary": summary,
                "n_rejected": len(self.rejected), "llm_calls": calls, "seconds": round(time.time() - t0, 1),
                "tools_used": dict(pd.Series([x["tool"] for x in self.log]).value_counts()) if self.log else {}}

    # ----------------------------------------------------------------------------------------------- tools
    def call(self, name: str, a: dict) -> Any:
        try:
            if name == "sql":
                return self._sql(a["query"])
            if name == "record_insight":
                return self._record(a.get("text", ""), a.get("evidence_sql", ""))
            if name == "finish":
                return "ok"
            if name.startswith("deep_") and self.cfg.deep:
                return self._deep(name, a)
            if name == "text_labels" and self.cfg.text:
                return self._text(a["column"], a.get("labels") or [])
            return f"unknown tool {name}"
        except Exception as e:
            return f"ERROR {type(e).__name__}: {str(e)[:300]}"

    def _sql(self, q: str, n: int = 20) -> str:
        res = self.con.execute(q).fetchdf()
        return f"{len(res)} rows; columns {list(res.columns)}\n{res.head(n).to_string(max_colwidth=60)}"

    def _record(self, text: str, sql: str) -> str:
        if not self.cfg.verify:
            self.ledger.append({"text": text, "sql": sql})
            return "recorded" + self._coverage(text)
        try:
            res = self.con.execute(sql).fetchdf()
        except Exception as e:
            self.rejected.append({"text": text, "why": "evidence SQL fails"})
            return f"REJECTED: evidence SQL fails ({str(e)[:150]})"
        shown = res.head(50).to_string(index=False)
        vals = [float(x) for x in re.findall(r"-?\d+(?:\.\d+)?(?:[eE]-?\d+)?", shown.replace(",", ""))]
        missing = [tok for tok in NUM.findall(text) if not number_supported(tok, vals)]
        if missing:
            self.rejected.append({"text": text, "why": f"numbers not in evidence: {missing}"})
            return f"REJECTED: these numbers are not in the evidence result: {missing}. Result was:\n{shown[:1500]}"
        if self.gli is not None:
            p = float(self.gli.scores([f"Evidence:\n{shown[:1500]}"], [text])[0, 0])
            if p < 0.1:
                self.rejected.append({"text": text, "why": f"not supported (score {p:.2f})"})
                return f"REJECTED: the evidence does not seem to support it (score {p:.2f}). Check and rephrase."
        self.ledger.append({"text": text, "sql": sql})
        return "ACCEPTED" + self._coverage(text)

    def _coverage(self, text: str) -> str:
        """Online signal: which analysis kinds the recorded insights cover so far, and which are still missing."""
        if not (self.cfg.coverage and self.gli is not None):
            return ""
        sc = self.gli.scores([text], KINDS)[0]
        self.covered |= {k for k, v in zip(KINDS, sc) if v > 0.5} or {KINDS[int(np.argmax(sc))]}
        missing = [k for k in KINDS if k not in self.covered]
        return (f". Covered so far: {', '.join(sorted(self.covered))}."
                + (f" Not yet covered: {', '.join(missing)}." if missing else " All kinds covered."))

    def _frame(self, exclude: list[str] | None = None) -> pd.DataFrame:
        """The frame the DEEP tool sees: date strings parsed to datetimes first; then long free text and unique string
        ids (which a tabular model cannot use) dropped."""
        d = self.df.drop(columns=[c for c in (exclude or []) if c in self.df.columns]).copy()

        def texty(c):  # object or pandas-3 string dtype
            return d[c].dtype == object or pd.api.types.is_string_dtype(d[c])

        for c in d.columns:
            if texty(c):
                parsed = pd.to_datetime(d[c], errors="coerce", format="mixed")
                if parsed.notna().mean() > 0.9:
                    d[c] = parsed
        keep = [c for c in d.columns if not (texty(c) and d[c].astype(str).str.len().mean() > 60)
                and not (texty(c) and d[c].nunique() == len(d))]
        return d[keep]

    def _deep(self, name: str, a: dict) -> str:
        t = a.get("target")
        if name == "deep_drivers":
            return json.dumps(self.deep.drivers(self._frame(a.get("exclude")), t, top=8))
        if name == "deep_what_if":
            return json.dumps(self.deep.what_if(self._frame(), t, a["column"]), default=str)
        if name == "deep_anomalies":
            fr = self._frame()
            return json.dumps(self.deep.anomalies(fr, t, id_col=a.get("id_column") if a.get("id_column") in fr else None, top=10),
                              default=str)
        if name == "deep_predict":
            tc = a.get("time_column")
            fr = self._frame()
            if tc and tc in self.df.columns:
                fr[tc] = self.df[tc]
            return json.dumps(self.deep.predict(fr, t, time_col=tc if tc in fr else None), default=str)
        if name == "deep_drift":
            tc = a["time_column"]
            ts = pd.to_datetime(self.df[tc], errors="coerce")
            cut = pd.to_datetime(a["cutoff"])
            fr = self._frame([tc])
            return json.dumps(self.deep.drift(fr[ts < cut], fr[ts >= cut]), default=str)
        return "unknown deep tool"

    def _text(self, col: str, labels: list[str]) -> str:
        labels = [str(x) for x in labels][:16]
        if not labels:
            return "give 3-16 label names"
        texts = self.df[col].astype(str).str[:1500].tolist()
        s = self.gli.scores(texts, labels)
        counts = {}
        for j, lab in enumerate(labels):
            name = f"{col}__{re.sub(r'[^0-9a-zA-Z]+', '_', lab).strip('_').lower()}"
            self.df[name] = (s[:, j] > 0.5).astype(int)
            counts[name] = int(self.df[name].sum())
        self.con.unregister("data")
        self.con.register("data", self.df)
        return f"added columns with counts: {counts}"


def g_eval_match(judge, pred: list[str], gold: list[str]) -> dict:
    """InsightBench-style one-to-many scoring with an open judge: for each gold insight, the best 1-10 match score over
    the predicted insights, averaged (scaled to 0-1). The judge is another LLM family than the analyst."""
    if not pred:
        return {"g_eval": 0.0, "per_gold": [0.0] * len(gold)}
    per = []
    for g in gold:
        listing = "\n".join(f"{i + 1}. {p}" for i, p in enumerate(pred))
        msgs = [{"role": "user", "content": f"Ground-truth insight: {g}\n\nPredicted insights:\n{listing}\n\n"
                 "Rate from 1 to 10 how well the BEST predicted insight matches the ground truth in meaning and in the "
                 "specific values. Answer with the number only."}]
        txt = judge.samples(msgs, n=1, temperature=0.0, max_tokens=8)[0].text
        m = re.search(r"\d+", txt or "")
        per.append((min(10, max(1, int(m.group(0)))) - 1) / 9 if m else 0.0)
    return {"g_eval": float(np.mean(per)), "per_gold": per}
