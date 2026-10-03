"""exp04 — which layer of a frozen relational foundation model gives the best entity embedding, on real RelBench tasks.

  python experiments/exp04_relbench_layers.py --tasks rel-f1/driver-dnf,rel-trial/study-outcome --modes random,label
Writes per task: layer x probe x {val AUROC, official test metrics}, the layer chosen on validation, and (label mode)
the model's own in-context prediction for comparison with the RelBench leaderboard's in-context entries.
"""
from __future__ import annotations

import argparse
import json
import os
import warnings
from pathlib import Path

import numpy as np
from tabfm_auto.logging_utils import RunLogger

from fer.relbench_layers import run_task

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="rel-f1/driver-dnf")
    ap.add_argument("--modes", default="random,label")
    ap.add_argument("--n-ctx", type=int, default=512); ap.add_argument("--n-train", type=int, default=4000)
    ap.add_argument("--k-children", type=int, default=50); ap.add_argument("--max-eval", type=int, default=20000)
    ap.add_argument("--device", default="cuda"); ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default="relbench_layers")
    ap.add_argument("--ctx-split", default="time", choices=["time", "random"],
                    help="time: probe-train rows strictly later than the context (default); random: the J18-J21 first pass")
    a = ap.parse_args()
    root = os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")
    for spec in a.tasks.split(","):
        dataset, task_name = spec.split("/")
        with RunLogger(f"{a.name}_{dataset}_{task_name}", {**vars(a), "task": spec}, root=root) as run:
            res = run_task(dataset, task_name, modes=tuple(a.modes.split(",")), n_ctx=a.n_ctx, n_train=a.n_train,
                           k_children=a.k_children, seed=a.seed, device=a.device, ctx_split=a.ctx_split, log=run, max_eval=a.max_eval)
            lines = [f"# {spec}: layer-wise probes of frozen Kumo Relational (val AUROC / official test metric)", ""]
            for mode, m in res["modes"].items():
                lines += [f"## context = {mode}", "", "| layer | linear val | linear test | kNN val | kNN test |", "|---|---|---|---|---|"]
                for L, r in m["layers"].items():
                    lt = r["linear"]["test"]; kt = r["knn"]["test"]
                    lines.append(f"| {L} | {r['linear']['val_auroc']:.4f} | {lt.get('roc_auc', np.nan):.4f} | "
                                 f"{r['knn']['val_auroc']:.4f} | {kt.get('roc_auc', np.nan):.4f} |")
                for p in ("linear", "knn"):
                    b = m[f"best_by_val_{p}"]
                    lines.append(f"\nbest layer by val ({p}): **{b['layer']}** -> test {b['test'].get('roc_auc', float('nan')):.4f}")
                if "icl_head" in m:
                    lines.append(f"model's own in-context prediction: val {m['icl_head']['val_auroc']:.4f}, "
                                 f"test {m['icl_head']['test'].get('roc_auc', float('nan')):.4f}")
                lines.append("")
            run.save_text("results.md", "\n".join(lines))
            json.dump(res, open(run.run_dir / "layers_rows.json", "w"), indent=1, default=float)
            run.finish({"layers": res})
            print("\n".join(lines), flush=True)


if __name__ == "__main__":
    main()
