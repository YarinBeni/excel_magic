"""DB agent: from a SQL database + a natural-language predictive question to a TabFM-Auto run.

Phase 1 (table building): a headless Claude Code session explores the database with ``tabfm-db`` and
writes ``build_table.sql`` that returns ONE ROW PER ENTITY with feature columns and the label column,
respecting the temporal cutoff. The harness materialises it, holds out a test split, and
Phase 2 runs the regular TabFM-Auto pipeline search around the frozen TFM on that table.

The same session is NOT reused for phase 2 so the pipeline-search agent never sees the raw DB.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from ..data.registry import TabularTask, infer_task_type
from ..logging_utils import RunLogger
from .claude_code import run_claude_code
from .search import run_search

PHASE1_TOOLS = ("Read,Write,Edit,Glob,Bash(tabfm-db:*),Bash(cat:*),Bash(ls:*),Bash(head:*),Bash(python:*),Bash(python3:*)")

PHASE1_RULES = """You are a data engineer agent. Build the training table for a predictive task from a SQL database.
Rules:
- Explore with `tabfm-db schema --db {db}` and `tabfm-db sql --db {db} --query "..."` (read-only SELECT).
- Write the final query to `build_table.sql`. It must return exactly one row per entity with:
  * an id column `{id_col}`,
  * feature columns computed ONLY from data dated on or before the cutoff ({cutoff}),
  * the label column `{target}` derived ONLY from data after the cutoff, as defined in the task.
  Features that peek past the cutoff are label leakage and invalidate the experiment.
- Prefer informative aggregates (counts, recency, sums over windows, ratios, category mixes) over raw ids.
  Keep <= 60 columns. Avoid free text. Categorical columns are fine as strings.
- Check your query with `tabfm-db materialize --db {db} --sql-file build_table.sql --target {target} --id-col {id_col} --out train.parquet`
  and fix errors until it succeeds. Then write FEATURES.md (what each feature means, in one line each) and stop.
- Never modify the database. Tables whose names start with "_" are off limits.
"""


def phase1_prompt(db: str, question: str, target: str, id_col: str, cutoff: str) -> str:
    return (PHASE1_RULES.format(db=db, target=target, id_col=id_col, cutoff=cutoff)
            + f"\n# Task\n{question}\n\nCutoff: {cutoff}. Label column: `{target}`. Entity id column: `{id_col}`.\n")


def run_db_agent(db_path: str, question: str, target: str, id_col: str, cutoff: str, llm_model: str = "sonnet",
                 phase1_max_turns: int = 40, phase1_minutes: int = 15, model_spec: str = "tabpfn:n_estimators=4",
                 search_kwargs: dict[str, Any] | None = None, name: str | None = None,
                 reference_task: TabularTask | None = None) -> dict[str, Any]:
    name = name or f"dbagent_{Path(db_path).stem}"
    cfg = {"db": db_path, "question": question, "target": target, "id_col": id_col, "cutoff": cutoff,
           "llm_model": llm_model, "model": model_spec, "search_kwargs": search_kwargs or {}}
    with RunLogger(name, cfg) as run:
        ws = run.run_dir / "phase1_workspace"
        ws.mkdir(exist_ok=True)
        db_abs = str(Path(db_path).resolve())
        prompt = phase1_prompt(db_abs, question, target, id_col, cutoff)
        (ws / "TASK.md").write_text(prompt)
        run.event("phase1_start", model=llm_model)
        info = run_claude_code(prompt, ws, run.run_dir / "phase1_stream.jsonl", model=llm_model,
                               max_turns=phase1_max_turns, timeout_s=phase1_minutes * 60, allowed_tools=PHASE1_TOOLS)
        run.event("phase1_end", **{k: v for k, v in info.items() if k != "result"})
        run.save_text("phase1_result.md", str(info.get("result") or ""))
        sql_path = ws / "build_table.sql"
        if not sql_path.exists():
            run.finish({"status_detail": "phase1 produced no build_table.sql", "phase1": info}, status="error")
            return {"error": "no build_table.sql"}
        # harness re-materialises the table itself (does not trust the agent's parquet)
        r = subprocess.run([sys.executable, "-m", "tabfm_auto.agent.db_tool", "materialize", "--db", db_abs,
                            "--sql-file", str(sql_path), "--target", target, "--id-col", id_col,
                            "--out", str(run.run_dir / "agent_table.parquet")], capture_output=True, text=True)
        run.save_text("materialize.log", r.stdout + "\n" + r.stderr)
        if r.returncode != 0:
            run.finish({"status_detail": "materialize failed", "stderr": r.stderr[-2000:]}, status="error")
            return {"error": "materialize failed", "stderr": r.stderr[-2000:]}
        df = pd.read_parquet(run.run_dir / "agent_table.parquet")
        y = df.pop(target)
        X = df.drop(columns=[id_col]) if id_col in df.columns else df
        ttype = infer_task_type(y)
        feats_md = (ws / "FEATURES.md").read_text() if (ws / "FEATURES.md").exists() else ""
        task = TabularTask(f"{name}_table", X, pd.Series(y.values, name=target), ttype,
                           {"description": question + "\n\nFeature table built by the DB agent from SQL. " + feats_md,
                            "source": f"db-agent:{db_path}", "columns": {c: "" for c in X.columns},
                            "build_table_sql": sql_path.read_text()})
        run.event("agent_table", **task.summary())
        shutil.copy(sql_path, run.run_dir / "build_table.sql")
        kw = dict(model_spec=model_spec, llm_model=llm_model, run_dir=run.run_dir / "phase2_search")
        kw.update(search_kwargs or {})
        run.event("phase2_start", **{k: v for k, v in kw.items() if k != "run_dir"})
        m2 = run_search(task, **kw)
        metrics: dict[str, Any] = {"phase1": {k: v for k, v in info.items() if k != "result"},
                                   "agent_table": task.summary(), "phase2": m2}
        if reference_task is not None:
            kw_ref = dict(kw)
            kw_ref["run_dir"] = run.run_dir / "reference_search"
            kw_ref["harness"] = "none"
            run.event("reference_start")
            metrics["reference_p0"] = run_search(reference_task, **kw_ref)
        run.finish(metrics)
        return metrics
