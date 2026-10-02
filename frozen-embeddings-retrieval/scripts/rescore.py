"""Re-score saved embeddings (runs/*/emb_<name>.npy) with the current benchmark code; writes rescored.json per run."""
import glob
import json
import sys
from pathlib import Path

import numpy as np
from tabfm_auto.data.synthetic_db import DEFAULT_DB

from fer.db import load_northwind, load_shop
from fer.retrieval_bench import run_benchmark

NORTHWIND = sys.argv[1] if len(sys.argv) > 1 else None
for d in sorted(set(glob.glob("runs/*bench_*") + glob.glob("runs/*retrieval*") + glob.glob("runs/*exp04*"))):
    cfg = json.load(open(f"{d}/config.json"))["config"]
    if cfg.get("kind") == "northwind":
        if not NORTHWIND:
            continue
        db = load_northwind(NORTHWIND)
    else:
        db = load_shop(cfg.get("db", str(DEFAULT_DB)))
    out = {}
    for f in sorted(Path(d).glob("emb_*.npy")):
        name = f.stem[4:]
        out[name] = run_benchmark(np.load(f), db, k_neighbors=cfg.get("k_neighbors", 20))
    json.dump({"db": db.name, "seed": cfg["seed"], "results": out}, open(f"{d}/rescored.json", "w"), indent=1)
    print(d, list(out))
