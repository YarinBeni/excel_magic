#!/bin/bash
#SBATCH --job-name=J12_relbench_churn
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=128G
#SBATCH --time=04:00:00
#SBATCH --output=sbatch/logs/J12_relbench_churn_%j.out
# J12 — segment-level probe on real data: RelBench rel-hm user-churn, kNN over frozen embeddings vs HGB / supervised TabPFN.
set -uo pipefail
echo "=== J12 relbench churn === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
python experiments/exp03_relbench_churn.py --n-customers 3000 --n-folds 3 --embedders agg,tabpfn_agg_kmeans --name J12_churn_smoke \
    || { echo "FAILED smoke"; cd ..; collect_and_push "J12: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
python experiments/exp03_relbench_churn.py --n-customers 20000 --n-folds 5 \
    --embedders row,agg,svd,svd_agg,tabpfn_agg_kmeans,tabpfn_agg_random,tabpfn_svd_kmeans --name J12_churn_full || echo "[J12] full FAILED"
cd ..
collect_and_push "J12: relbench user-churn probe (job ${SLURM_JOB_ID})"
echo "=== J12 done $(date) ==="
