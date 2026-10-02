"""Run the TabFM-Auto paper protocol on TabArena v0.1 datasets (needs api.openml.org + the `openml` package).

Small-CPU smoke test (3 smallest datasets, Lite = split r0f0 only, tiny budget):
  python examples/05_tabarena_protocol.py --max-instances 1000 --lite --budget-evals 6 --budget-minutes 15

Paper-like run on an H100 node (all 51 datasets, all official splits, 6h budget per dataset, Opus):
  python examples/05_tabarena_protocol.py --all --model kumo-tabular-s:n_estimators=8 --llm opus \
      --budget-evals 96 --budget-minutes 360
"""
from __future__ import annotations

import argparse
import warnings

from tabfm_auto.benchmarks.tabarena import list_datasets, run_tabarena

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--datasets", default=None, help="comma-separated TabArena dataset names")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--max-instances", type=int, default=2500)
    ap.add_argument("--max-features", type=int, default=200)
    ap.add_argument("--lite", action="store_true", help="score split r0f0 only (TabArena-Lite)")
    ap.add_argument("--model", default="tabpfn:n_estimators=8")
    ap.add_argument("--harness", default="claude-code", choices=["claude-code", "openai", "cli", "heuristic", "none"])
    ap.add_argument("--llm-base-url", default=None); ap.add_argument("--agent-cmd", default=None)
    ap.add_argument("--llm", default="sonnet")
    ap.add_argument("--budget-evals", type=int, default=24)
    ap.add_argument("--budget-minutes", type=int, default=120)
    ap.add_argument("--max-rows", type=int, default=10000)
    ap.add_argument("--name", default="tabarena")
    ap.add_argument("--cv-repeats", default="1", help="repeats of the 3-fold judge; 'auto' = 3 when n_train < 1000")
    ap.add_argument("--select", default="best", help="final pick: best | gated1 | gated2 (paired per-fold gate vs P0)")
    a = ap.parse_args()
    if a.datasets:
        wanted = set(a.datasets.split(","))
        ds = [d for d in list_datasets() if d.name in wanted]
    elif a.all:
        ds = list_datasets()
    else:
        ds = list_datasets(max_instances=a.max_instances, max_features=a.max_features)
    print(f"{len(ds)} datasets:", [d.name for d in ds])
    res = run_tabarena(ds, model_spec=a.model, lite=a.lite, harness=a.harness, llm_model=a.llm, budget_evals=a.budget_evals,
                       budget_minutes=a.budget_minutes, max_rows=a.max_rows, name=a.name, llm_base_url=a.llm_base_url,
                       agent_cmd=a.agent_cmd, cv_repeats=a.cv_repeats if a.cv_repeats == "auto" else int(a.cv_repeats), select_rule=a.select)
    print(res.groupby(["dataset", "pipeline"]).error.mean().unstack())


if __name__ == "__main__":
    main()
