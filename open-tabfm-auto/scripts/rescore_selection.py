"""Re-score finished TabArena searches under different final-selection rules, without re-running any agent.

For each tabarena run dir (with search/<dataset>/workspace/{evals.jsonl,candidates/}), apply the rules in
tabfm_auto.harness.selection to the recorded per-fold CV scores, then score each rule's pick (or ensemble) on every
official split. Writes rescored.json + selection_rules.md into the run dir.

  python scripts/rescore_selection.py --runs artifacts/runs/*J2_heuristic_* --rules best,gated1,gated2,ens3
"""
from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

import pandas as pd

from tabfm_auto.benchmarks.tabarena import list_datasets, load_openml_task, official_split
from tabfm_auto.harness.evaluator import evaluate_holdout
from tabfm_auto.harness.selection import evaluate_holdout_ensemble, select


def rescore_run(run_dir: Path, rules: list[str], lite: bool = False) -> pd.DataFrame | None:
    cfg = json.loads((run_dir / "config.json").read_text())
    model_spec = cfg.get("model", "tabpfn"); max_rows = cfg.get("max_rows", 10000)
    ds_by_name = {d.name: d for d in list_datasets()}
    rows = []
    for sdir in sorted((run_dir / "search").glob("*")):
        ws = sdir / "workspace"
        if not (ws / "evals.jsonl").exists() or sdir.name not in ds_by_name:
            continue
        d = ds_by_name[sdir.name]
        evals = [json.loads(l) for l in (ws / "evals.jsonl").read_text().splitlines() if l.strip()]
        picks = {rule: select(evals, rule) for rule in rules}
        task, oml = load_openml_task(d)
        t0 = time.time()
        for (r, f) in d.splits(lite):
            tr, te = official_split(oml, r, f)
            Xtr, ytr = task.X.iloc[tr].reset_index(drop=True), task.y.iloc[tr].reset_index(drop=True)
            Xte, yte = task.X.iloc[te].reset_index(drop=True), task.y.iloc[te].reset_index(drop=True)
            cache: dict[tuple[str, ...], float | None] = {}
            for rule, names in picks.items():
                key = tuple(names)
                if key not in cache:
                    paths = [ws / "candidates" / n for n in names]
                    if len(paths) == 1:
                        res = evaluate_holdout(paths[0], Xtr, ytr, Xte, yte, task.task_type, model_spec, seed=r * 3 + f, max_rows=max_rows)
                    else:
                        res = evaluate_holdout_ensemble(paths, Xtr, ytr, Xte, yte, task.task_type, model_spec, seed=r * 3 + f, max_rows=max_rows)
                    cache[key] = res.get("score")
                rows.append({"dataset": d.name, "repeat": r, "fold": f, "rule": rule, "pick": "+".join(names), "error": cache[key]})
        print(f"[rescore] {run_dir.name} {d.name}: picks={ {k: '+'.join(v) for k, v in picks.items()} } {time.time() - t0:.0f}s", flush=True)
    if not rows:
        return None
    df = pd.DataFrame(rows)
    piv = df.pivot_table(index="dataset", columns="rule", values="error", aggfunc="mean")
    out = {"run": run_dir.name, "rules": rules, "per_split": rows}
    (run_dir / "rescored.json").write_text(json.dumps(out, default=float))
    md = piv.to_markdown(floatfmt=".5f")
    (run_dir / "selection_rules.md").write_text(md + "\n")
    df.to_csv(run_dir / "selection_rules.csv", index=False)
    return df


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--runs", nargs="+", required=True)
    ap.add_argument("--rules", default="best,gated1,gated2,ens3")
    ap.add_argument("--lite", action="store_true")
    a = ap.parse_args()
    rules = a.rules.split(",")
    if "p0" not in rules:
        rules = ["p0"] + rules
    for r in a.runs:
        rescore_run(Path(r), rules, a.lite)


if __name__ == "__main__":
    main()
