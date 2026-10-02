"""Exp02: TabFM-Auto search. A headless Claude Code agent evolves pipeline.py around frozen TabPFN v2.

Example: python examples/02_pipeline_search.py --dataset synth_physics --budget-evals 10 --budget-minutes 20
         python examples/02_pipeline_search.py --dataset wine --harness none   # P0 only, no LLM
"""
from __future__ import annotations

import argparse
import warnings

from tabfm_auto.agent.search import run_search
from tabfm_auto.data import load_task

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dataset", required=True)
    ap.add_argument("--model", default="tabpfn:n_estimators=4")
    ap.add_argument("--harness", default="claude-code", choices=["claude-code", "none"])
    ap.add_argument("--llm-model", default="sonnet")
    ap.add_argument("--budget-evals", type=int, default=12)
    ap.add_argument("--budget-minutes", type=int, default=30)
    ap.add_argument("--max-turns", type=int, default=80)
    ap.add_argument("--test-size", type=float, default=0.3)
    ap.add_argument("--n-folds", type=int, default=3)
    ap.add_argument("--max-rows", type=int, default=10000)
    ap.add_argument("--eval-timeout", type=int, default=900)
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--baselines", default="hgb,rf")
    ap.add_argument("--name", default=None)
    a = ap.parse_args()
    task = load_task(a.dataset)
    m = run_search(task, model_spec=a.model, harness=a.harness, llm_model=a.llm_model, budget_evals=a.budget_evals,
                   budget_minutes=a.budget_minutes, max_turns=a.max_turns, test_size=a.test_size, n_folds=a.n_folds,
                   seed=a.seed, max_rows=a.max_rows, eval_timeout_s=a.eval_timeout,
                   baselines=tuple(b for b in a.baselines.split(",") if b), name=a.name)
    keys = ["dataset", "metric", "p0_cv", "best_cv", "p0_test", "best_test", "test_improvement_pct", "n_evals",
            "baseline_hgb_test", "baseline_rf_test"]
    print({k: m.get(k) for k in keys})


if __name__ == "__main__":
    main()
