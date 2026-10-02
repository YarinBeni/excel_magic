#!/bin/bash
#SBATCH --job-name=J7d_relbench_hm_svd
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=128G
#SBATCH --time=06:00:00
#SBATCH --output=sbatch/logs/J7d_relbench_hm_svd_%j.out
# J7d — rel-hm again with the interaction signal as TFM input: SVD(64) of the purchase matrix (+ aggregates) through the frozen TabPFN vs the raw factors.
# Rows: GlobalPopularity, PastVisit (reproduce 0.30 / 0.89 test), then our training-free kNN-CF over customer embeddings
# (row-only, hand aggregates, frozen TabPFN hidden state with k-means / random target) and PastVisit+kNN hybrids.
set -uo pipefail
echo "=== J7d relbench hm === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
# smoke on 3000 queries first (fails fast on API drift), then the full val+test
python experiments/exp02_relbench_hm.py --splits val --sample-queries 3000 --embedders svd,tabpfn_svd_kmeans \
    --name J7d_hm_smoke || { echo "FAILED smoke"; cd ..; collect_and_push "J7d: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
python experiments/exp02_relbench_hm.py --splits val,test --embedders svd,svd_agg,tabpfn_svd_kmeans,tabpfn_svd_random,agg,tabpfn_agg_kmeans \
    --hist-days "${HIST_DAYS:-365}" --k-neighbors "${KNN:-50}" --name J7d_hm_svd_full || echo "[J7] full run FAILED"
cd ..
collect_and_push "J7d: relbench rel-hm results (job ${SLURM_JOB_ID})"
echo "=== J7d done $(date) ==="
