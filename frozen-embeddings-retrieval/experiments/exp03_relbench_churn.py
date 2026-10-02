"""exp03 — segment-level probe on RelBench rel-hm user-churn: frozen-TFM embeddings + kNN vs supervised references.

  python experiments/exp03_relbench_churn.py --embedders row,agg,svd,tabpfn_agg_kmeans,tabpfn_agg_random,tabpfn_svd_kmeans
"""
from __future__ import annotations

import argparse
import json
import os
import warnings
from pathlib import Path

from tabfm_auto.logging_utils import RunLogger

from fer.relbench_churn import run_probe

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--embedders", default="row,agg,svd,svd_agg,tabpfn_agg_kmeans,tabpfn_agg_random,tabpfn_svd_kmeans")
    ap.add_argument("--hist-days", type=int, default=365); ap.add_argument("--k-neighbors", type=int, default=50)
    ap.add_argument("--n-customers", type=int, default=20000); ap.add_argument("--n-folds", type=int, default=5)
    ap.add_argument("--device", default="cuda"); ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default="relbench_churn")
    a = ap.parse_args()
    with RunLogger(a.name, vars(a), root=os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")) as run:
        res = run_probe(a.embedders.split(","), a.hist_days, a.n_customers or None, a.k_neighbors, a.n_folds, a.seed,
                        a.device, log=run)
        lines = ["| method | AUROC |", "|---|---|"]
        for n, r in res["rows"].items():
            lines.append(f"| {n} | " + ("ERR" if "error" in r else f"{r['auroc']:.4f}") + " |")
        lines.append(f"\ncustomers={res['n_customers']} at {res['t0']}, pos_rate={res['pos_rate']:.3f}, hist_days={a.hist_days}, "
                     f"k={a.k_neighbors}, {a.n_folds}-fold over customers (not the official temporal test split)")
        run.save_text("results.md", "\n".join(lines))
        json.dump(res, open(run.run_dir / "churn_rows.json", "w"), indent=1, default=float)
        run.finish({"churn": res})
        print("\n".join(lines))


if __name__ == "__main__":
    main()
