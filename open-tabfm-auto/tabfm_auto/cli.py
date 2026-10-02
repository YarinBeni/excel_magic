"""``tabfm-auto``: one entry point for the three things people do with this library.

  tabfm-auto search --dataset breast_cancer --model tabpfn --llm sonnet --budget-evals 12
  tabfm-auto search --csv data.csv --target label --description "..." --model tabpfn
  tabfm-auto db     --db shop.sqlite --question "..." --target churned --id-col customer_id --cutoff 2025-07-01
  tabfm-auto models                      # which backbones are runnable; download with `tabfm-models download <name>`
  tabfm-auto runs                        # summary table of runs/
"""
from __future__ import annotations

import argparse
import sys
import warnings

warnings.filterwarnings("ignore")


def _task_from_args(a):
    from .data import load_task

    if a.csv:
        import pandas as pd

        from .data.registry import TabularTask, infer_task_type

        df = pd.read_csv(a.csv)
        y = df.pop(a.target)
        ttype = a.task_type or infer_task_type(y)
        if ttype != "regression":
            y = pd.Series(pd.factorize(y)[0], name=a.target)
        return TabularTask(a.name or a.csv.rsplit("/", 1)[-1].split(".")[0], df, y, ttype,
                           {"description": a.description or "", "source": a.csv, "columns": {c: "" for c in df.columns}})
    return load_task(a.dataset)


def cmd_search(a) -> int:
    from .agent.search import run_search

    task = _task_from_args(a)
    m = run_search(task, model_spec=a.model, harness=a.harness, llm_model=a.llm, budget_evals=a.budget_evals,
                   budget_minutes=a.budget_minutes, max_turns=a.max_turns, test_size=a.test_size, n_folds=a.n_folds,
                   seed=a.seed, max_rows=a.max_rows, baselines=tuple(b for b in a.baselines.split(",") if b), name=a.run_name,
                   llm_base_url=a.llm_base_url)
    keys = ["dataset", "metric", "p0_cv", "best_cv", "p0_test", "best_test", "test_improvement_pct", "n_evals"]
    print({k: m.get(k) for k in keys})
    return 0


def cmd_db(a) -> int:
    from .agent.db_agent import run_db_agent

    m = run_db_agent(a.db, a.question, target=a.target, id_col=a.id_col, cutoff=a.cutoff, llm_model=a.llm,
                     model_spec=a.model, phase1_minutes=a.phase1_minutes,
                     search_kwargs={"budget_evals": a.budget_evals, "budget_minutes": a.budget_minutes,
                                    "max_rows": a.max_rows, "baselines": ("hgb",)}, name=a.run_name)
    p2 = m.get("phase2", {})
    print({k: p2.get(k) for k in ["p0_cv", "best_cv", "p0_test", "best_test", "n_evals"]})
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="tabfm-auto", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("--model", default="tabpfn:n_estimators=4", help="frozen backbone spec (see `tabfm-auto models`)")
        p.add_argument("--llm", default="sonnet", help="LLM for the coding agent: a Claude alias for --harness claude-code, "
                                                      "or a model id served by an OpenAI-compatible endpoint for --harness openai")
        p.add_argument("--llm-base-url", default=None, help="OpenAI-compatible endpoint (vLLM, SGLang, Ollama); or $OPENAI_BASE_URL")
        p.add_argument("--budget-evals", type=int, default=12)
        p.add_argument("--budget-minutes", type=int, default=30)
        p.add_argument("--max-rows", type=int, default=10000)
        p.add_argument("--run-name", default=None)

    s = sub.add_parser("search", help="TabFM-Auto pipeline search on a dataset")
    src = s.add_mutually_exclusive_group(required=True)
    src.add_argument("--dataset", help="registry name (see tabfm_auto.data.list_tasks) or openml:<name>")
    src.add_argument("--csv", help="a CSV file; needs --target")
    s.add_argument("--target"); s.add_argument("--task-type", choices=["binary", "multiclass", "regression"])
    s.add_argument("--description", default=None, help="what the table is about (helps the agent)")
    s.add_argument("--name", default=None)
    s.add_argument("--harness", default="claude-code", choices=["claude-code", "openai", "heuristic", "none"],
                   help="claude-code (headless Claude Code), openai (any OpenAI-compatible LLM), heuristic (no LLM), none (P0 only)")
    s.add_argument("--max-turns", type=int, default=80); s.add_argument("--test-size", type=float, default=0.3)
    s.add_argument("--n-folds", type=int, default=3); s.add_argument("--seed", type=int, default=0)
    s.add_argument("--baselines", default="hgb")
    common(s); s.set_defaults(fn=cmd_search)

    d = sub.add_parser("db", help="build the training table from a SQL database with an agent, then search")
    d.add_argument("--db", required=True, help="sqlite path or SQLAlchemy URL")
    d.add_argument("--question", required=True); d.add_argument("--target", required=True)
    d.add_argument("--id-col", required=True); d.add_argument("--cutoff", required=True)
    d.add_argument("--phase1-minutes", type=int, default=15)
    common(d); d.set_defaults(fn=cmd_db)

    m = sub.add_parser("models", help="which backbones are runnable here")
    m.set_defaults(fn=lambda a: __import__("tabfm_auto.models.weights", fromlist=["main"]).main(["status"]))
    r = sub.add_parser("runs", help="summarize runs/")
    r.set_defaults(fn=lambda a: __import__("tabfm_auto.logging_utils", fromlist=["summarize_cli"]).summarize_cli([]))

    a = ap.parse_args(argv)
    return a.fn(a) or 0


if __name__ == "__main__":
    sys.exit(main())
