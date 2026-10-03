#!/bin/bash
#SBATCH --job-name=J22_kumo_own_pred
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=01:00:00
#SBATCH --output=sbatch/logs/J22_kumo_own_pred_%j.out
# J22 — Kumo Relational's own prediction is at chance on rel-event/user-repeat (0.49) while a probe on its graph layer
# reaches 0.79. Print the raw per-class output columns vs the extracted score, with rel-f1/driver-dnf as the control.
set -uo pipefail
echo "=== J22 kumo own-prediction check === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
mkdir -p reports/checks
( cd frozen-embeddings-retrieval && python scripts/check_kumo_own_pred.py --tasks "${TASKS:-rel-f1/driver-dnf,rel-event/user-repeat,rel-event/user-ignore}" ) \
    2>&1 | grep -v "Fetching" | tee reports/checks/kumo_own_pred_${SLURM_JOB_ID}.txt
collect_and_push "J22: Kumo Relational own-prediction check (job ${SLURM_JOB_ID})"
echo "=== J22 done $(date) ==="
