"""Prompts for the pipeline-search coding agent (TabFM-Auto, Section 3.2)."""
from __future__ import annotations

import json
from typing import Any

SYSTEM_RULES = """You are the TabFM-Auto pipeline-search agent. A FROZEN tabular foundation model is fixed; you only
edit the data pipeline around it (pipeline.py) to minimise the 3-fold cross-validation error reported by
`tabfm-eval`. Rules:
- Only edit pipeline.py. Only run `tabfm-eval` (and read-only shell commands / small python snippets to inspect
  train.parquet). Never try to read, list or guess anything outside this directory. Never use the network.
- Never replace the frozen model, train another model inside pipeline.py, or hard-code label information.
- Keep every improvement that lowers the CV score; revert edits that make it worse (candidates/ keeps snapshots).
- You may use pandas, numpy, scipy and scikit-learn (transforms only) inside pipeline.py.
- Think about the column names and task description: domain formulas, ratios, log transforms of skewed targets,
  missing-value sentinels, entity co-occurrence counts for ID columns, class-balanced context views, prior /
  temperature calibration in postprocess().
- When the evaluation budget is exhausted or you cannot improve further, stop. The harness picks the best
  candidate automatically; leave pipeline.py equal to your best candidate.
"""


def task_brief(task: dict[str, Any], metadata: dict[str, Any], p0: dict[str, Any] | None) -> str:
    cols = metadata.get("columns", {})
    col_lines = "\n".join(f"  - {c}: {d}" if d else f"  - {c}" for c, d in cols.items())
    p0_line = (f"Identity pipeline P0 scored {p0['score']:.5f} ({p0.get('metric')}, lower is better) "
               f"with {p0.get('n_features')} features." if p0 and p0.get("status") == "ok"
               else f"Identity pipeline P0 FAILED: {p0.get('error') if p0 else 'not run'}")
    return f"""# Task
Dataset: {task['dataset']}  |  type: {task['task_type']}  |  target column: `{task['target']}`
Metric: {task['metric']} (lower is better), 3-fold CV on train.parquet ({task['n_train']} rows, {task['n_cols']} columns).
Frozen model: {task['model_spec']} (TabPFN v2, in-context learner; context limited to max_rows={task['max_rows']}).
Budget: {task['budget_evals']} evaluations / {task['budget_minutes']} minutes. {p0_line}

## Dataset description
{metadata.get('description', '')}

## Columns
{col_lines}

{('## Extra notes' + chr(10) + metadata['long_description'][:2500]) if metadata.get('long_description') else ''}

# How to work
1. `cat pipeline.py` to see the current (identity) pipeline and the contract of the four hooks.
2. Inspect the data briefly (`python -c "import pandas as pd; df=pd.read_parquet('train.parquet'); print(df.describe(include='all').T)"`).
3. Edit pipeline.py, then run `tabfm-eval` and read the score. Repeat.
4. Finish with pipeline.py equal to your best candidate and write 3-6 lines in NOTES.md explaining what helped.
"""


def workspace_task_md(task: dict[str, Any], metadata: dict[str, Any], p0: dict[str, Any] | None) -> str:
    return SYSTEM_RULES + "\n" + task_brief(task, metadata, p0)


def json_block(obj: Any) -> str:
    return "```json\n" + json.dumps(obj, indent=2, default=str) + "\n```"
