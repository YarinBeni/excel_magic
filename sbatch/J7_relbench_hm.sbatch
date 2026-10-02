#!/bin/bash
#SBATCH --job-name=J7_relbench_hm
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=128G
#SBATCH --time=06:00:00
#SBATCH --output=sbatch/logs/J7_relbench_hm_%j.out
# J7 — the published benchmark: RelBench rel-hm / user-item-purchase (MAP@12), official evaluator.
# Rows: GlobalPopularity, PastVisit (reproduce 0.30 / 0.89 test), then our training-free kNN-CF over customer embeddings
# (row-only, hand aggregates, frozen TabPFN hidden state with k-means / random target) and PastVisit+kNN hybrids.
set -uo pipefail
echo "=== J7 relbench hm === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
# smoke on 3000 queries first (fails fast on API drift), then the full val+test
python experiments/exp02_relbench_hm.py --splits val --sample-queries 3000 --embedders row,agg,tabpfn_agg_kmeans \
    --name J7_hm_smoke || { echo "FAILED smoke"; cd ..; collect_and_push "J7: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
python experiments/exp02_relbench_hm.py --splits val,test --embedders row,agg,tabpfn_agg_kmeans,tabpfn_agg_random,tabpfn_row_kmeans \
    --hist-days "${HIST_DAYS:-365}" --k-neighbors "${KNN:-50}" --name J7_hm_full || echo "[J7] full run FAILED"
cd ..
collect_and_push "J7: relbench rel-hm results (job ${SLURM_JOB_ID})"
echo "=== J7 done $(date) ==="
