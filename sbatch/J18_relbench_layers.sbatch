#!/bin/bash
#SBATCH --job-name=J18_relbench_layers
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=07:00:00
#SBATCH --output=sbatch/logs/J18_relbench_layers_%j.out
# J18 — which layer of frozen Kumo Relational gives the best entity embedding, on REAL RelBench entity-classification
# tasks (official splits, official test evaluator). Layers: pre-GNN row embedding, after GNN, each of 12 ICL layers,
# final. Context: random target (task-agnostic embedding) and true labels (the model's own in-context prediction).
# Smallest tasks first; results pushed after each task.
set -uo pipefail
echo "=== J18 relbench layers === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
# smoke: tiny run, and a variance check that every tapped layer actually differs across rows (catches hook bugs)
python - <<'PY' || { echo "FAILED smoke"; cd ..; collect_and_push "J18: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
import numpy as np, warnings; warnings.filterwarnings("ignore")
import relbench
from sdm.models import KumoRelational
from fer.relbench_layers import Subgrapher, embed_rows
task = relbench.load_dataset("rel-f1").load_task("driver-dnf"); db = task.get_db(upto_test_timestamp=False)
tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
print("entity", task.entity_table, task.entity_col, "time", task.time_col, "target", task.target_col, "train rows", len(tr))
sub = Subgrapher(db, task.entity_table); print("children", sub.children)
ctx = tr.sample(64, random_state=0).drop_duplicates(task.entity_col).reset_index(drop=True)
rows = tr.sample(200, random_state=1).reset_index(drop=True)
m = KumoRelational(device="cuda"); m.eval()
E, sc = embed_rows(m, sub, ctx, ctx[task.target_col].to_numpy(), rows, task.entity_col, task.time_col, "cuda", batch=100)
for k, v in E.items():
    print(f"{k:8s} dim {v.shape[1]} row-std {float(v.std(0).mean()):.4f} unique {len(np.unique(v.round(4), axis=0))}")
print("icl scores range", float(sc.min()), float(sc.max()))
assert all(len(np.unique(v.round(4), axis=0)) > 10 for v in E.values()), "a tapped layer is constant across rows"
PY
TASKS="${TASKS:-rel-f1/driver-dnf,rel-f1/driver-top3,rel-trial/study-outcome,rel-event/user-repeat,rel-event/user-ignore,rel-avito/user-visits,rel-avito/user-clicks,rel-hm/user-churn}"
IFS=',' read -ra TL <<< "$TASKS"
for T in "${TL[@]}"; do
    echo "--- $T $(date)"
    python experiments/exp04_relbench_layers.py --tasks "$T" --modes random,label --name J18_layers || echo "[J18] $T FAILED"
    ( cd .. && collect_and_push "J18: layer sweep $T (job ${SLURM_JOB_ID})" )
done
echo "=== J18 done $(date) ==="
