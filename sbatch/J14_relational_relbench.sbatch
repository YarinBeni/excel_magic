#!/bin/bash
#SBATCH --job-name=J14_relational_relbench
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=08:00:00
#SBATCH --output=sbatch/logs/J14_relational_relbench_%j.out
# J14 — the relational foundation model on real data: frozen Kumo Relational (NVIDIA sdm) customer embeddings built from
# the rel-hm tables (customers <- transactions -> articles), k-means / random in-context target.
#   a) entity level: user-churn kNN probe next to TabPFN and the supervised references (J12 protocol)
#   b) item level: user-item-purchase kNN-CF rows with the official evaluator (J7 protocol)
set -uo pipefail
echo "=== J14 relational relbench === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
python experiments/exp03_relbench_churn.py --n-customers 2000 --n-folds 3 --embedders agg,kumo_relational_kmeans --name J14_kumo_smoke \
    || { echo "FAILED smoke"; cd ..; collect_and_push "J14: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
python experiments/exp03_relbench_churn.py --n-customers 20000 --n-folds 5 \
    --embedders agg,tabpfn_agg_kmeans,kumo_relational_kmeans,kumo_relational_random --name J14_kumo_churn || echo "[J14] churn FAILED"
cd ..; collect_and_push "J14: Kumo Relational churn probe (job ${SLURM_JOB_ID})"; cd frozen-embeddings-retrieval
python experiments/exp02_relbench_hm.py --splits val,test --embedders agg,kumo_relational_kmeans \
    --hist-days 365 --k-neighbors 50 --name J14_kumo_hm_full || echo "[J14] item-level FAILED"
cd ..
collect_and_push "J14: Kumo Relational rel-hm item-level rows (job ${SLURM_JOB_ID})"
echo "=== J14 done $(date) ==="
