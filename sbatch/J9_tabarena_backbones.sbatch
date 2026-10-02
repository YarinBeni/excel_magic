#!/bin/bash
#SBATCH --job-name=J9_tabarena_backbones
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=14:00:00
#SBATCH --array=0-1
#SBATCH --output=sbatch/logs/J9_tabarena_backbones_%A_%a.out
# J9 — per-backbone search (paper 5.4 conclusion): the TabArena wave-1 protocol with the no-LLM heuristic (currently the
# best free searcher on TabArena) around the two strongest open backbones: Kumo Tabular-L (task 0) and TabICLv2 (task 1).
# Compare with J2 (Kumo-S) and the paper's TabFM / TabFM-Auto rows.
set -uo pipefail
echo "=== J9 backbones task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
MODELS=("kumo-tabular-l:n_estimators=8,device=cuda" "tabicl:device=cuda"); TAGS=("kumoL" "tabiclv2")
MODEL="${MODELS[$SLURM_ARRAY_TASK_ID]}"; TAG="${TAGS[$SLURM_ARRAY_TASK_ID]}"
cd open-tabfm-auto
mapfile -t DATASETS < <(python -c "from tabfm_auto.benchmarks.tabarena import list_datasets; print('\n'.join(d.name for d in list_datasets(max_instances=2500, max_features=200)))")
for DS in "${DATASETS[@]}"; do
    echo "--- $DS $(date)"
    python examples/05_tabarena_protocol.py --datasets "$DS" --model "$MODEL" --harness heuristic \
        --budget-evals 24 --budget-minutes 90 --name "J9_heuristic_${TAG}_${DS}" || echo "[J9] $DS FAILED"
    ( cd .. && collect_and_push "J9: tabarena heuristic $TAG $DS (job ${SLURM_JOB_ID})" )
done
echo "=== J9 $TAG done $(date) ==="
