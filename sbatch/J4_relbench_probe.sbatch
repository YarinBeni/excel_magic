#!/bin/bash
#SBATCH --job-name=J4_relbench
#SBATCH --partition=voltagepark
#SBATCH --cpus-per-task=16
#SBATCH --mem=96G
#SBATCH --time=04:00:00
#SBATCH --output=sbatch/logs/J4_relbench_%j.out
# J4 — probe the published benchmark: download RelBench rel-hm, print the schema/sizes, compute the two trivial published
# baselines (global popularity, past visit) for user-item-purchase with RelBench's own evaluator (target: 0.30 / 0.89 MAP@12 x100)
# so the adapter in fer/relbench_adapter.py can be finished against real API output.
set -uo pipefail
echo "=== J4 relbench probe === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
python - <<'PY' 2>&1 | tee reports/J4_relbench_probe.md
import traceback, numpy as np, pandas as pd
print("# RelBench rel-hm probe\n")
try:
    import relbench
    print("relbench", getattr(relbench, "__version__", "?"))
    ds = relbench.load_dataset("rel-hm")
    print("dataset:", ds)
    task = ds.load_task("user-item-purchase") if hasattr(ds, "load_task") else relbench.tasks.get_task("rel-hm", "user-item-purchase", download=True)
    print("task:", task, "\n")
    db = task.get_db(upto_test_timestamp=False)
    for name, t in db.table_dict.items():
        print(f"- {name}: {len(t.df)} rows, cols={list(t.df.columns)[:12]}, pkey={t.pkey_col}, fkeys={t.fkey_col_to_pkey_table}, time={t.time_col}")
    for split in ("train", "val", "test"):
        tbl = task.get_table(split, mask_input_cols=False)
        print(f"- task table {split}: {len(tbl.df)} rows, cols={list(tbl.df.columns)}")
    # trivial baselines, val split, with the official evaluator
    tx = db.table_dict["transactions"].df
    val = task.get_table("val", mask_input_cols=False).df
    K = getattr(task, "eval_k", 12)
    src, dst = task.src_entity_col, task.dst_entity_col
    ts_col = task.time_col if hasattr(task, "time_col") else "timestamp"
    cutoff = val[ts_col].min()
    past = tx[tx[db.table_dict["transactions"].time_col] < cutoff]
    pop = past[task.dst_entity_table[:-1] + "_id" if False else [c for c in past.columns if c.endswith("article_id")][0]].value_counts().index[:K].to_numpy()
    pred_pop = np.tile(pop, (len(val), 1))
    print("\nGlobalPopularity val:", task.evaluate(pred_pop, task.get_table("val")))
    col_src = [c for c in past.columns if c.endswith("customer_id")][0]; col_dst = [c for c in past.columns if c.endswith("article_id")][0]
    last = past.sort_values(db.table_dict["transactions"].time_col).groupby(col_src)[col_dst].apply(lambda s: list(dict.fromkeys(s[::-1]))[:K])
    pv = np.array([ (last.get(u, []) + list(pop))[:K] for u in val[src] ])
    print("PastVisit val:", task.evaluate(pv, task.get_table("val")))
except Exception:
    traceback.print_exc()
    print("\nPROBE FAILED: see traceback; adjust fer/relbench_adapter.py to this API")
PY
commit_push_sandbox "J4: relbench probe (job ${SLURM_JOB_ID})" reports
echo "=== J4 done $(date) ==="
