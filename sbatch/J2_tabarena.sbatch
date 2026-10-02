#!/bin/bash
#SBATCH --job-name=J2_tabarena
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=06:00:00
#SBATCH --array=0-16
#SBATCH --output=sbatch/logs/J2_tabarena_%A_%a.out
# J2 — the paper's TabArena protocol on the 17 datasets with <= 2500 rows, one dataset per array task:
# search on fold-0 training split with the no-LLM heuristic (env HARNESS overrides), frozen Kumo Tabular-S (env MODEL),
# score P0 and P* on all official splits, compare with the paper's Table 9.
set -uo pipefail
echo "=== J2 tabarena task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd open-tabfm-auto
DS=$(python -c "from tabfm_auto.benchmarks.tabarena import list_datasets; print(list_datasets(max_instances=2500, max_features=200)[${SLURM_ARRAY_TASK_ID}].name)")
MODEL="${MODEL:-kumo-tabular-s:n_estimators=8,device=cuda}"
HARNESS="${HARNESS:-heuristic}"
echo "dataset=$DS model=$MODEL harness=$HARNESS"
python examples/05_tabarena_protocol.py --datasets "$DS" --model "$MODEL" --harness "$HARNESS" --llm "${LLM:-sonnet}" \
    --budget-evals "${BUDGET_EVALS:-24}" --budget-minutes "${BUDGET_MINUTES:-90}" --name "J2_${HARNESS}_${DS}" || echo "[J2] $DS FAILED"
cd ..
collect_and_push "J2: tabarena $DS $HARNESS (job ${SLURM_JOB_ID})"
echo "=== J2 $DS done $(date) ==="
