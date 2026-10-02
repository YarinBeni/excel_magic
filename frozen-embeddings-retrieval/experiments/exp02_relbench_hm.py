"""Exp02: RelBench rel-hm / user-item-purchase with frozen-embedding kNN retrieval, official evaluator.

  python experiments/exp02_relbench_hm.py --splits val,test --embedders row,agg,tabpfn_agg_kmeans,tabpfn_agg_random
"""
from __future__ import annotations

import argparse
import json
import os
import warnings
from pathlib import Path

from tabfm_auto.logging_utils import RunLogger

from fer.relbench_hm import load_task, run_split

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dataset", default="rel-hm"); ap.add_argument("--task", default="user-item-purchase")
    ap.add_argument("--splits", default="val,test")
    ap.add_argument("--embedders", default="row,agg,tabpfn_agg_kmeans,tabpfn_agg_random")
    ap.add_argument("--hist-days", type=int, default=365); ap.add_argument("--k-neighbors", type=int, default=50)
    ap.add_argument("--sample-queries", type=int, default=None, help="subsample query customers (smoke test)")
    ap.add_argument("--device", default="cuda"); ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default="relbench_hm")
    a = ap.parse_args()
    with RunLogger(a.name, vars(a), root=os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")) as run:
        task, db = load_task(a.dataset, a.task)
        out = {}
        for split in a.splits.split(","):
            out[split] = run_split(task, db, split, a.embedders.split(","), a.hist_days, a.k_neighbors, a.seed, a.device,
                                   log=run, sample_queries=a.sample_queries)
        lines = ["| method | " + " | ".join(f"{s} MAP@K x100" for s in out) + " |", "|---|" + "---|" * len(out)]
        names = list(dict.fromkeys(n for s in out.values() for n in s["rows"]))
        for n in names:
            cells = []
            for s in out.values():
                r = s["rows"].get(n, {})
                cells.append("ERR" if "error" in r or "map" not in r else f"{100 * r['map']:.3f}")
            lines.append(f"| {n} | " + " | ".join(cells) + " |")
        lines.append("\nqueries: " + ", ".join(f"{s}={v['n_queries']}" for s, v in out.items()) +
                     f"; hist_days={a.hist_days}; k_neighbors={a.k_neighbors}; published test rows (x100): GlobalPop 0.30, "
                     "PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93")
        run.save_text("results.md", "\n".join(lines))
        run.finish({"splits": out})
        print("\n".join(lines))
        json.dump(out, open(run.run_dir / "relbench_rows.json", "w"), indent=1, default=float)


if __name__ == "__main__":
    main()
