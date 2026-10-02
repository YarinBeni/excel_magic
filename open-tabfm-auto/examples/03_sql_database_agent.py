"""Exp03: DB agent. Phase 1 builds a feature table from SQL (headless Claude Code + tabfm-db),
phase 2 runs the TabFM-Auto search around frozen TabPFN on that table. A hand-written reference
feature table (tabfm_auto.data.synthetic_db.load_db_task) is scored with P0 for comparison.

Example: python examples/03_sql_database_agent.py --budget-evals 8 --budget-minutes 20
"""
from __future__ import annotations

import argparse
import warnings

from tabfm_auto.agent.db_agent import run_db_agent
from tabfm_auto.data.synthetic_db import DB_TASK_DESCRIPTION, DEFAULT_DB, generate, load_db_task

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=str(DEFAULT_DB))
    ap.add_argument("--cutoff", default="2025-07-01")
    ap.add_argument("--llm-model", default="sonnet")
    ap.add_argument("--model", default="tabpfn:n_estimators=4")
    ap.add_argument("--budget-evals", type=int, default=8)
    ap.add_argument("--budget-minutes", type=int, default=20)
    ap.add_argument("--phase1-minutes", type=int, default=15)
    ap.add_argument("--max-rows", type=int, default=10000)
    ap.add_argument("--no-reference", action="store_true")
    ap.add_argument("--name", default="exp03_db_agent")
    a = ap.parse_args()
    generate(a.db)
    question = DB_TASK_DESCRIPTION.format(cutoff=a.cutoff)
    ref = None if a.no_reference else load_db_task(path=a.db)
    m = run_db_agent(a.db, question, target="churned", id_col="customer_id", cutoff=a.cutoff, llm_model=a.llm_model,
                     model_spec=a.model, phase1_minutes=a.phase1_minutes,
                     search_kwargs={"budget_evals": a.budget_evals, "budget_minutes": a.budget_minutes,
                                    "max_rows": a.max_rows, "baselines": ("hgb",)},
                     name=a.name, reference_task=ref)
    p2 = m.get("phase2", {})
    print({k: p2.get(k) for k in ["p0_cv", "best_cv", "p0_test", "best_test", "n_evals", "baseline_hgb_test"]})
    if m.get("reference_p0"):
        print("reference hand-built table P0 test:", m["reference_p0"].get("p0_test"))


if __name__ == "__main__":
    main()
