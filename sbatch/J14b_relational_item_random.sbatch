#!/bin/bash
#SBATCH --job-name=J14b_relational_item_random
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=04:00:00
#SBATCH --output=sbatch/logs/J14b_relational_item_random_%j.out
# J14 — the relational foundation model on real data: frozen Kumo Relational (NVIDIA sdm) customer embeddings built from
# the rel-hm tables (customers <- transactions -> articles), k-means / random in-context target.
#   a) entity level: user-churn kNN probe next to TabPFN and the supervised references (J12 protocol)
#   b) item level: user-item-purchase kNN-CF rows with the official evaluator (J7 protocol)
set -uo pipefail
echo "=== J14b relational relbench === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
python experiments/exp02_relbench_hm.py --splits val,test --embedders agg,kumo_relational_random \
    --hist-days 365 --k-neighbors 50 --name J14b_kumo_random_hm_full || echo "[J14] item-level FAILED"
cd ..
collect_and_push "J14b: Kumo Relational (random target) rel-hm item-level rows (job ${SLURM_JOB_ID})"
echo "=== J14b done $(date) ==="

