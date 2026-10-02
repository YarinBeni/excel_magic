#!/bin/bash
#SBATCH --job-name=J11_judge
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=10:00:00
#SBATCH --array=0-1
#SBATCH --output=sbatch/logs/J11_judge_%A_%a.out
# J11 — the judge on small tables: repeated 3-fold CV (3 repeats = 9 folds) at the standard budget, on the datasets with
# < 1000 training rows (blood, diabetes, QSAR_fish, credit-g, anneal) + the three bundled tasks. Compare with J2/J8/J6
# (1 repeat). Task 0 heuristic, task 1 pi + Qwen3-Coder.
set -uo pipefail
echo "=== J11 judge task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh; source sbatch/_agents.sh
PORT=$((8500 + SLURM_ARRAY_TASK_ID)); BASE="http://localhost:$PORT/v1"; LLM="-"
if [ "$SLURM_ARRAY_TASK_ID" = "0" ]; then TAG="heuristic"; HARNESS_ARGS=(--harness heuristic); SARGS=(--harness heuristic)
else
    LLM="Qwen/Qwen3-Coder-30B-A3B-Instruct"; TAG="pi_qwen3coder"
    start_vllm "qwen3_coder hermes" || { echo "FAILED: vllm"; exit 1; }; trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
    setup_pi "$BASE" "$LLM" || exit 1
    HARNESS_ARGS=(--harness cli --agent-cmd pi --llm "$LLM" --llm-base-url "$BASE"); SARGS=("${HARNESS_ARGS[@]}")
fi
export OPENAI_BASE_URL="$BASE" OPENAI_API_BASE="$BASE" OPENAI_API_KEY=dummy OPENAI_MODEL="$LLM"
cd open-tabfm-auto
for ds in synth_physics synth_entities breast_cancer; do
    tabfm-auto search --dataset $ds "${SARGS[@]}" --model "kumo-tabular-s:n_estimators=4,device=cuda" --cv-repeats 3 \
        --budget-evals 16 --budget-minutes 40 --baselines hgb --run-name "J11_${TAG}_cv3_$ds" || echo "[J11] $ds FAILED"
done
( cd .. && collect_and_push "J11: judge cv3 $TAG synthetic (job ${SLURM_JOB_ID})" )
for DS in blood-transfusion-service-center diabetes QSAR_fish_toxicity credit-g anneal; do
    echo "--- $DS $(date)"
    python examples/05_tabarena_protocol.py --datasets "$DS" --model "kumo-tabular-s:n_estimators=8,device=cuda" "${HARNESS_ARGS[@]}" \
        --budget-evals 24 --budget-minutes 90 --cv-repeats 3 --name "J11_${TAG}_cv3_${DS}" || echo "[J11] $DS FAILED"
    ( cd .. && collect_and_push "J11: judge cv3 $TAG $DS (job ${SLURM_JOB_ID})" )
done
echo "=== J11 $TAG done $(date) ==="
