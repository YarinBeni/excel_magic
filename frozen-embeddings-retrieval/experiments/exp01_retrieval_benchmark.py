"""Exp01: are relational-model hidden states useful graph-aware entity embeddings?

Compares customer embeddings (row-only, hand aggregates, from-scratch GNN, pre-trained OpenRFM hidden
state, random-weight OpenRFM) on (A) latent-segment retrieval and (B) future-purchase kNN retrieval.

Example: python experiments/exp01_retrieval_benchmark.py --embedders row,agg,gnn,openrfm,openrfm_random
"""
from __future__ import annotations

import argparse
import os
import time
import warnings
from pathlib import Path

import numpy as np
from tabfm_auto.data.synthetic_db import DEFAULT_DB, generate
from tabfm_auto.logging_utils import RunLogger

from fer.db import load_northwind, load_shop
from fer.embedders import EMBEDDERS
from fer.retrieval_bench import run_benchmark

warnings.filterwarnings("ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--db", default=str(DEFAULT_DB))
    ap.add_argument("--kind", default="synth", choices=["synth", "northwind"])
    ap.add_argument("--embedders", default="row,agg,gnn,openrfm,openrfm_ctx16,openrfm_random")
    ap.add_argument("--k-neighbors", type=int, default=20)
    ap.add_argument("--seed", type=int, default=0)
    ap.add_argument("--name", default="exp01_retrieval_benchmark")
    a = ap.parse_args()
    cfg = vars(a)
    with RunLogger(a.name, cfg, root=os.environ.get("TABFM_RUNS_ROOT", Path(__file__).resolve().parents[1] / "runs")) as run:
        if a.kind == "synth":
            generate(a.db)
            db = load_shop(a.db)
        else:
            db = load_northwind(a.db)
        run.event("db", name=db.name, n_customers=len(db.customers), n_products=len(db.products),
                  n_past_items=len(db.past_items()), n_future_items=len(db.future_items()), cutoff=str(db.cutoff.date()))
        rows = []
        for name in a.embedders.split(","):
            t0 = time.time()
            try:
                E = EMBEDDERS[name](db, seed=a.seed, log=run)
                np.save(run.run_dir / f"emb_{name}.npy", E)
                r = run_benchmark(E, db, k_neighbors=a.k_neighbors)
                r.update({"embedder": name, "seconds": round(time.time() - t0, 1), "status": "ok"})
            except Exception as e:
                run.log.exception("embedder %s failed", name)
                r = {"embedder": name, "status": "error", "error": f"{type(e).__name__}: {e}",
                     "seconds": round(time.time() - t0, 1)}
            rows.append(r)
            run.event("result", **r)
            run.info("%-16s %s", name, {k: v for k, v in r.items() if k not in ("embedder",)})
        lines = ["| embedder | dim | seg P@10 | seg MAP@10 | seg kNN acc | fut MAP@10 | fut recall@10 | fut hit@10 | s |",
                 "|---|---|---|---|---|---|---|---|---|"]
        for r in rows:
            if r["status"] != "ok":
                lines.append(f"| {r['embedder']} | ERR | | | | | | | {r['seconds']} |")
                continue
            s = r.get("segment", {})
            f = r["future_purchase"]
            lines.append(f"| {r['embedder']} | {r['dim']} | {s.get('P@10', float('nan')):.3f} | {s.get('MAP@10', float('nan')):.3f} | "
                         f"{s.get('kNN_acc', float('nan')):.3f} | {f['MAP@10']:.3f} | {f['recall@10']:.3f} | {f['hit@10']:.3f} | {r['seconds']} |")
        if rows and rows[0].get("future_purchase"):
            f = rows[0]["future_purchase"]
            lines.append(f"| popularity baseline | - | {rows[0].get('segment', {}).get('chance_P@10', float('nan')):.3f} (chance) | | | "
                         f"{f['pop_MAP@10']:.3f} | {f['pop_recall@10']:.3f} | | |")
        run.save_text("results.md", "\n".join(lines))
        run.finish({"db": db.name, "results": rows, "embedders": a.embedders})
        print("\n".join(lines))


if __name__ == "__main__":
    main()
