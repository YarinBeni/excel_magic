#!/bin/bash
#SBATCH --job-name=J17_recommended
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=10:00:00
#SBATCH --array=0-1
#SBATCH --output=sbatch/logs/J17_recommended_%A_%a.out
# J17 — the recommended free configuration, end to end on TabArena wave 1: the strongest open backbone (Kumo Tabular-L),
# repeated-CV judge on small tables, 1-s.e. gated pick, and a 20% acceptance slice of the training split that the
# candidate must beat P0 on. Task 0: no-LLM heuristic. Task 1: pi + Qwen3-Coder-30B (vLLM). Compare with J9 (Kumo-L, plain).
set -uo pipefail
echo "=== J17 recommended task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh; source sbatch/_agents.sh
PORT=$((8700 + SLURM_ARRAY_TASK_ID)); BASE="http://localhost:$PORT/v1"; LLM="-"
if [ "$SLURM_ARRAY_TASK_ID" = "0" ]; then TAG="heuristic"; HARNESS_ARGS=(--harness heuristic)
else
    LLM="Qwen/Qwen3-Coder-30B-A3B-Instruct"; TAG="pi_qwen3coder"
    start_vllm "qwen3_coder hermes" || { echo "FAILED: vllm"; exit 1; }; trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
    setup_pi "$BASE" "$LLM" || exit 1; HARNESS_ARGS=(--harness cli --agent-cmd pi --llm "$LLM" --llm-base-url "$BASE")
fi
export OPENAI_BASE_URL="$BASE" OPENAI_API_BASE="$BASE" OPENAI_API_KEY=dummy OPENAI_MODEL="$LLM"
cd open-tabfm-auto
mapfile -t DATASETS < <(python -c "from tabfm_auto.benchmarks.tabarena import list_datasets; print('\n'.join(d.name for d in list_datasets(max_instances=2500, max_features=200)))")
for DS in "${DATASETS[@]}"; do
    echo "--- $DS $(date)"
    python examples/05_tabarena_protocol.py --datasets "$DS" --model "kumo-tabular-l:n_estimators=8,device=cuda" "${HARNESS_ARGS[@]}" \
        --budget-evals 24 --budget-minutes 60 --cv-repeats auto --select gated1 --accept-holdout 0.2 \
        --name "J17_${TAG}_kumoL_rec_${DS}" || echo "[J17] $DS FAILED"
    ( cd .. && collect_and_push "J17: recommended config $TAG $DS (job ${SLURM_JOB_ID})" )
done
echo "=== J17 $TAG done $(date) ==="
