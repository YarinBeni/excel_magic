"""TabArena v0.1 protocol, as run in the TabFM-Auto paper (Section 4.1 / Appendix B.1).

For each dataset: run the pipeline search ONCE, on the training split of repeat 0 / fold 0 (3-fold inner CV,
test rows never visible), freeze the best pipeline P*, then score P0 and P* on every official split
(9 or 30 per dataset; ``lite=True`` scores split r0f0 only, TabArena-Lite style). Results are written in a
long format (dataset, split, method, error) so Elo against any pool can be computed with
``tabfm_auto.benchmarks.elo``. The paper's per-dataset numbers (Table 9: TabFM and TabFM-Auto) ship with
the package for a direct comparison.

Needs the ``openml`` package and network access to api.openml.org (datasets + official split indices are cached
locally by openml). Set ``OPENML_CACHE_DIR`` to relocate the cache (e.g. a shared cluster filesystem).
"""
from __future__ import annotations

import csv
import json
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from ..agent.search import run_search
from ..data.registry import TabularTask
from ..harness.evaluator import evaluate_holdout
from ..logging_utils import RunLogger

DATA = Path(__file__).resolve().parent / "data"


@dataclass(frozen=True)
class TabArenaDataset:
    name: str
    tid: int
    problem_type: str  # binary | multiclass | regression
    metric: str        # roc_auc | log_loss | rmse (TabArena names)
    n_instances: int
    n_features: int
    n_splits: int      # 30 = 10 repeats x 3 folds, 9 = 3 x 3

    @property
    def n_repeats(self) -> int:
        return self.n_splits // 3

    def splits(self, lite: bool = False) -> list[tuple[int, int]]:
        if lite:
            return [(0, 0)]
        return [(r, f) for r in range(self.n_repeats) for f in range(3)]


def list_datasets(max_instances: int | None = None, max_features: int | None = None,
                  problem_types: tuple[str, ...] | None = None) -> list[TabArenaDataset]:
    out = []
    with open(DATA / "tabarena_v01_datasets.csv") as f:
        for r in csv.DictReader(f):
            d = TabArenaDataset(r["dataset"], int(r["tid"]), r["problem_type"], r["metric"], int(r["n_instances"]),
                                int(r["n_features"]), int(r["n_splits"]))
            if max_instances and d.n_instances > max_instances:
                continue
            if max_features and d.n_features > max_features:
                continue
            if problem_types and d.problem_type not in problem_types:
                continue
            out.append(d)
    return out


def paper_table9() -> pd.DataFrame:
    """Per-dataset official test scores from arXiv 2609.37989 Table 9 (TabFM and two TabFM-Auto configurations)."""
    return pd.read_csv(DATA / "paper_table9.csv")


def load_openml_task(d: TabArenaDataset, cache_dir: str | Path | None = None):
    """Returns (TabularTask, openml task). Official split indices come from ``task.get_train_test_split_indices``."""
    import openml

    if cache_dir or os.environ.get("OPENML_CACHE_DIR"):
        openml.config.set_root_cache_directory(str(cache_dir or os.environ["OPENML_CACHE_DIR"]))
    task = openml.tasks.get_task(d.tid, download_splits=True, download_data=True, download_qualities=False,
                                 download_features_meta_data=True)
    X, y, _, _ = task.get_dataset().get_data(target=task.target_name)
    X = pd.DataFrame(X)
    y = pd.Series(y, name=task.target_name)
    if d.problem_type == "regression":
        y = y.astype(float)
    else:
        y = pd.Series(pd.factorize(y)[0], name=task.target_name, index=y.index)
    desc = (task.get_dataset().description or "")[:4000]
    meta = {"description": f"TabArena dataset {d.name} (OpenML task {d.tid}). {desc}", "source": f"openml:{d.tid}",
            "columns": {c: "" for c in X.columns}, "target_name": task.target_name}
    return TabularTask(d.name, X.reset_index(drop=True), y.reset_index(drop=True), d.problem_type, meta), task


def official_split(task, repeat: int, fold: int) -> tuple[np.ndarray, np.ndarray]:
    tr, te = task.get_train_test_split_indices(fold=fold, repeat=repeat)
    return np.asarray(tr), np.asarray(te)


def run_tabarena(datasets: list[TabArenaDataset], model_spec: str = "tabpfn:n_estimators=8", lite: bool = True,
                 harness: str = "claude-code", llm_model: str = "sonnet", budget_evals: int = 24, budget_minutes: int = 120,
                 max_turns: int = 150, max_rows: int = 10000, eval_timeout_s: int = 1800, baselines: tuple[str, ...] = (),
                 name: str = "tabarena", cache_dir: str | Path | None = None, run_root: Path | None = None,
                 llm_base_url: str | None = None, agent_cmd: str | None = None, cv_repeats: int | str = 1,
                 select_rule: str = "best") -> pd.DataFrame:
    """Paper protocol over ``datasets``. Returns the long results frame (also saved as results.csv)."""
    cfg = {"datasets": [d.name for d in datasets], "model": model_spec, "lite": lite, "harness": harness,
           "llm_model": llm_model, "budget_evals": budget_evals, "budget_minutes": budget_minutes, "max_rows": max_rows,
           "cv_repeats": cv_repeats, "select_rule": select_rule}
    rows: list[dict[str, Any]] = []
    with RunLogger(name, cfg, root=run_root) as run:
        paper = paper_table9().set_index("dataset")
        for d in datasets:
            t0 = time.time()
            try:
                task, oml = load_openml_task(d, cache_dir)
            except Exception as e:
                run.warning("%s: could not load from OpenML (%s)", d.name, e)
                run.event("dataset_load_failed", dataset=d.name, error=str(e)[:500])
                continue
            tr0, te0 = official_split(oml, 0, 0)
            run.event("dataset", **task.summary(), tid=d.tid, n_train_r0f0=len(tr0))
            # 1. search on r0f0's training split (the agent never sees te0)
            m = run_search(task, model_spec=model_spec, harness=harness, llm_model=llm_model, budget_evals=budget_evals,
                           budget_minutes=budget_minutes, max_turns=max_turns, max_rows=max_rows,
                           eval_timeout_s=eval_timeout_s, baselines=baselines, split=(tr0, te0),
                           name=f"{name}_{d.name}", run_dir=run.run_dir / "search" / d.name,
                           llm_base_url=llm_base_url, agent_cmd=agent_cmd, cv_repeats=cv_repeats, select_rule=select_rule)
            sdir = run.run_dir / "search" / d.name
            p0 = sdir / "workspace" / "candidates" / "eval_001.py"
            pstar = sdir / "best_pipeline.py" if (sdir / "best_pipeline.py").exists() else p0
            # 2. score frozen P0 and P* on every official split
            for (r, f) in d.splits(lite):
                tr, te = official_split(oml, r, f)
                for method, pipe in (("P0", p0), ("P*", pstar)):
                    res = evaluate_holdout(pipe, task.X.iloc[tr].reset_index(drop=True), task.y.iloc[tr].reset_index(drop=True),
                                           task.X.iloc[te].reset_index(drop=True), task.y.iloc[te].reset_index(drop=True),
                                           task.task_type, model_spec, seed=r * 3 + f, max_rows=max_rows, log=run)
                    row = {"dataset": d.name, "tid": d.tid, "repeat": r, "fold": f, "task": f"{d.name}/r{r}f{f}",
                           "method": f"{model_spec.split(':')[0]}+{method}", "pipeline": method, "metric": res["metric"],
                           "error": res.get("score"), "status": res["status"], "err_msg": res.get("error"),
                           "seconds": res["elapsed_s"]}
                    rows.append(row)
                    run.event("split_result", **row)
            # 3. per-dataset summary vs the paper
            df = pd.DataFrame([x for x in rows if x["dataset"] == d.name and x["status"] == "ok"])
            summ = df.groupby("pipeline").error.mean().to_dict() if len(df) else {}
            pap = paper.loc[d.name].to_dict() if d.name in paper.index else {}
            run.info("%-40s P0=%.5f P*=%.5f | paper TabFM=%s TabFM-Auto(Opus5)=%s | %.0fs", d.name,
                     summ.get("P0", float("nan")), summ.get("P*", float("nan")), pap.get("tabfm"), pap.get("tabfm_auto_opus5"),
                     time.time() - t0)
            run.event("dataset_summary", dataset=d.name, ours=summ, paper=pap, search_metrics={k: m.get(k) for k in
                      ("p0_cv", "best_cv", "n_evals", "best_candidate")})
            pd.DataFrame(rows).to_csv(run.run_dir / "results.csv", index=False)
        results = pd.DataFrame(rows)
        results.to_csv(run.run_dir / "results.csv", index=False)
        comp = compare_to_paper(results)
        comp.to_csv(run.run_dir / "comparison_to_paper.csv", index=False)
        run.save_text("comparison_to_paper.md", comp.to_markdown(index=False, floatfmt=".4f") if len(comp) else "no results")
        run.finish({"n_datasets": results.dataset.nunique() if len(results) else 0, "n_rows": len(results),
                    "comparison": comp.to_dict("records")})
        return results


def compare_to_paper(results: pd.DataFrame) -> pd.DataFrame:
    """Per-dataset mean error of P0 / P* next to the paper's TabFM / TabFM-Auto numbers and relative gains."""
    if not len(results):
        return pd.DataFrame()
    ok = results[results.status == "ok"]
    piv = ok.pivot_table(index="dataset", columns="pipeline", values="error", aggfunc="mean")
    paper = paper_table9().set_index("dataset")
    out = piv.join(paper, how="left").reset_index()
    out["our_gain_%"] = 100 * (out["P0"] - out["P*"]) / out["P0"]
    out["paper_gain_opus5_%"] = 100 * (out["tabfm"] - out["tabfm_auto_opus5"]) / out["tabfm"]
    out["ours_vs_paper_tabfm"] = out["P*"] / out["tabfm"]
    return out


def geometric_mean_error(results: pd.DataFrame, pipeline: str = "P*") -> float:
    ok = results[(results.status == "ok") & (results.pipeline == pipeline)]
    per_ds = ok.groupby("dataset").error.mean()
    return float(np.exp(np.log(per_ds.clip(lower=1e-12)).mean()))


def write_elo_input(results: pd.DataFrame, path: str | Path) -> None:
    """Long (method, task, error) rows ready to be concatenated with a leaderboard pool for ``elo_ratings``."""
    results[results.status == "ok"][["method", "task", "error"]].to_csv(path, index=False)
    json.dumps({"note": "concatenate with the pool's per-split errors, then elo_ratings(df, anchor='RandomForest (default)')"})
