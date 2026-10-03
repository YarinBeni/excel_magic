"""exp06 — every open tabular foundation model, every layer, on real RelBench entity tasks: supervised probes, label-free
geometry, and CKA within and across models.

  python experiments/exp06_model_layers.py --tasks rel-f1/driver-dnf --models tabpfn,tabpfn-2.5,tabiclv2,kumo-s,kumo-m,kumo-l
"""
from __future__ import annotations

import argparse
import json
import os
import warnings
from pathlib import Path

import numpy as np
from tabfm_auto.logging_utils import RunLogger

from fer.model_layers import run_task

warnings.filterwarnings("ignore")
SPECS = {"tabpfn": "tabpfn:n_estimators=1,device=cuda,inference_precision=float32",
         "tabpfn-2.5": "tabpfn-2.5:n_estimators=1,device=cuda,inference_precision=float32",
         "tabiclv2": "tabiclv2-sdm:n_estimators=1,device=cuda",
         "kumo-s": "kumo-tabular-s:n_estimators=1,device=cuda",
         "kumo-m": "kumo-tabular-m:n_estimators=1,device=cuda",
         "kumo-l": "kumo-tabular-l:n_estimators=1,device=cuda"}


def summary(res) -> str:
    lines = [f"# {res['dataset']}/{res['task']}: every model, every layer (test AUROC, official evaluator)", "",
             "| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |",
             "|---|---|---|---|---|---|---|---|"]
    for mname, per in res["models"].items():
        for target, m in per.items():
            if "error" in m:
                lines.append(f"| {mname} | {target} | ERROR {m['error'][:80]} | | | | | |")
                continue
            L = m["layers"]
            last = [k for k in L if k != "raw_features"][-1]
            gv = {k: v["spearman_with_linear_test"] for k, v in m["geometry_vs_probe"].items()}
            gbest = max(gv, key=lambda k: abs(np.nan_to_num(gv[k]))) if gv else "-"
            own = m.get("icl_head", {}).get("test", {}).get("roc_auc")
            lines.append(f"| {mname} | {target} | {L['raw_features']['linear']['test_auroc']:.4f} | {m['best_by_val_linear']['layer']} | "
                         f"{m['best_by_val_linear']['test_auroc']:.4f} | {L[last]['linear']['test_auroc']:.4f} | "
                         f"{'' if own is None else f'{own:.4f}'} | {gbest} ({gv.get(gbest, float('nan')):+.2f}) |")
    return "\n".join(lines)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="rel-f1/driver-dnf")
    ap.add_argument("--models", default=",".join(SPECS))
    ap.add_argument("--targets", default="random,kmeans,label")
    ap.add_argument("--n-ctx", type=int, default=3000); ap.add_argument("--n-train", type=int, default=4000)
    ap.add_argument("--max-eval", type=int, default=5000); ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default="model_layers")
    ap.add_argument("--ctx-split", default="time", choices=["time", "random"],
                    help="time: probe-train rows strictly later than the context (default); random: the J18-J21 first pass")
    a = ap.parse_args()
    specs = {m: SPECS[m] for m in a.models.split(",")}
    root = os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")
    for spec in a.tasks.split(","):
        dataset, task_name = spec.split("/")
        with RunLogger(f"{a.name}_{dataset}_{task_name}", {**vars(a), "task": spec}, root=root) as run:
            res = run_task(dataset, task_name, specs, targets=tuple(a.targets.split(",")), n_ctx=a.n_ctx,
                           n_train=a.n_train, seed=a.seed, max_eval=a.max_eval, ctx_split=a.ctx_split, log=run)
            md = summary(res)
            run.save_text("results.md", md)
            json.dump(res, open(run.run_dir / "layers_rows.json", "w"), indent=1, default=float)
            run.finish({"model_layers": {k: v for k, v in res.items() if k != "models"}})
            print(md, flush=True)


if __name__ == "__main__":
    main()
