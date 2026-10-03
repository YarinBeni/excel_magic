#!/bin/bash
#SBATCH --job-name=J19_relbench_tab_layers
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=07:00:00
#SBATCH --output=sbatch/logs/J19_relbench_tab_layers_%j.out
# J19 — the single-table counterpart of J18: frozen TabPFN v2 over per-entity features (entity attributes + linked-row
# counts / recency / means, same temporal window), every transformer block, four in-context targets (zeros, random,
# k-means, true labels), on the same 8 real RelBench entity tasks with the official test evaluator.
set -uo pipefail
echo "=== J19 relbench tabular layers === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
TASKS="${TASKS:-rel-f1/driver-dnf,rel-f1/driver-top3,rel-trial/study-outcome,rel-event/user-repeat,rel-event/user-ignore,rel-avito/user-visits,rel-avito/user-clicks,rel-hm/user-churn}"
IFS=',' read -ra TL <<< "$TASKS"
for T in "${TL[@]}"; do
    echo "--- $T $(date)"
    python experiments/exp05_relbench_tabular_layers.py --tasks "$T" --name J19_tab_layers || echo "[J19] $T FAILED"
    ( cd .. && collect_and_push "J19: tabular layer sweep $T (job ${SLURM_JOB_ID})" )
done
echo "=== J19 done $(date) ==="
