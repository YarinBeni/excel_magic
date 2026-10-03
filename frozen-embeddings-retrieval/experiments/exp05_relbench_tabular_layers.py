"""exp05 — layer-wise row embeddings of frozen TabPFN v2 on the same real RelBench entity tasks as exp04 (matched
comparison with the relational model), with four in-context targets (zeros, random, k-means, true labels).

  python experiments/exp05_relbench_tabular_layers.py --tasks rel-f1/driver-dnf
"""
from __future__ import annotations

import argparse
import json
import os
import warnings
from pathlib import Path

from tabfm_auto.logging_utils import RunLogger

from fer.relbench_tabular_layers import run_task

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="rel-f1/driver-dnf")
    ap.add_argument("--targets", default="zeros,random,kmeans,label")
    ap.add_argument("--n-ctx", type=int, default=3000); ap.add_argument("--n-train", type=int, default=4000)
    ap.add_argument("--max-eval", type=int, default=20000); ap.add_argument("--device", default="cuda")
    ap.add_argument("--seed", type=int, default=0); ap.add_argument("--name", default="relbench_tab_layers")
    ap.add_argument("--ctx-split", default="time", choices=["time", "random"],
                    help="time: probe-train rows strictly later than the context (default); random: the J18-J21 first pass")
    a = ap.parse_args()
    root = os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")
    for spec in a.tasks.split(","):
        dataset, task_name = spec.split("/")
        with RunLogger(f"{a.name}_{dataset}_{task_name}", {**vars(a), "task": spec}, root=root) as run:
            res = run_task(dataset, task_name, targets=tuple(a.targets.split(",")), n_ctx=a.n_ctx, n_train=a.n_train,
                           seed=a.seed, device=a.device, ctx_split=a.ctx_split, log=run, max_eval=a.max_eval)
            lines = [f"# {spec}: layer-wise probes of frozen TabPFN v2 over per-entity features ({res['n_features']} features)", ""]
            for target, m in res["targets"].items():
                if "error" in m:
                    lines += [f"## target = {target}: ERROR {m['error']}", ""]
                    continue
                lines += [f"## target = {target}", "", "| layer | linear val | linear test | kNN val | kNN test |", "|---|---|---|---|---|"]
                for L, r in m["layers"].items():
                    lines.append(f"| {L} | {r['linear']['val_auroc']:.4f} | {r['linear']['test'].get('roc_auc', float('nan')):.4f} | "
                                 f"{r['knn']['val_auroc']:.4f} | {r['knn']['test'].get('roc_auc', float('nan')):.4f} |")
                for p in ("linear", "knn"):
                    b = m[f"best_by_val_{p}"]
                    lines.append(f"\nbest block by val ({p}): **{b['layer']}** -> test {b['test'].get('roc_auc', float('nan')):.4f}")
                if "icl_head" in m:
                    lines.append(f"model's own in-context prediction: test {m['icl_head']['test'].get('roc_auc', float('nan')):.4f}")
                lines.append("")
            run.save_text("results.md", "\n".join(lines))
            json.dump(res, open(run.run_dir / "layers_rows.json", "w"), indent=1, default=float)
            run.finish({"tabular_layers": res})
            print("\n".join(lines), flush=True)


if __name__ == "__main__":
    main()
