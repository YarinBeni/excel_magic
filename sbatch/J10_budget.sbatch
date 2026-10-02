#!/bin/bash
#SBATCH --job-name=J10_budget
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=16:00:00
#SBATCH --array=0-2
#SBATCH --output=sbatch/logs/J10_budget_%A_%a.out
# J10 — budget curve (paper Fig. 4) on the five datasets where the paper gains most (anneal +17.6%, Marketing_Campaign
# +15.8%, airfoil +14.6%, hazelnut +8.7%, diabetes +6.9%): 64 evals / 150 min instead of 16-24 / 40-90, judge with
# repeated CV on the small ones (--cv-repeats auto). Task 0 heuristic, 1 pi + Qwen3-Coder, 2 tool loop + GLM-4.5-Air-FP8.
set -uo pipefail
echo "=== J10 budget task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh; source sbatch/_agents.sh
DATASETS=(anneal Marketing_Campaign airfoil_self_noise hazelnut-spread-contaminant-detection diabetes)
PORT=$((8400 + SLURM_ARRAY_TASK_ID)); BASE="http://localhost:$PORT/v1"; LLM="-"
case "$SLURM_ARRAY_TASK_ID" in
  0) TAG="heuristic"; HARNESS_ARGS=(--harness heuristic) ;;
  1) LLM="Qwen/Qwen3-Coder-30B-A3B-Instruct"; TAG="pi_qwen3coder"
     start_vllm "qwen3_coder hermes" || { echo "FAILED: vllm"; exit 1; }; trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
     setup_pi "$BASE" "$LLM" || exit 1; HARNESS_ARGS=(--harness cli --agent-cmd pi --llm "$LLM" --llm-base-url "$BASE") ;;
  2) LLM="zai-org/GLM-4.5-Air-FP8"; TAG="openai_glm45air"; export VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=65536
     start_vllm "glm45 hermes" || { echo "FAILED: vllm"; exit 1; }; trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
     HARNESS_ARGS=(--harness openai --llm "$LLM" --llm-base-url "$BASE") ;;
esac
export OPENAI_BASE_URL="$BASE" OPENAI_API_BASE="$BASE" OPENAI_API_KEY=dummy OPENAI_MODEL="$LLM"
cd open-tabfm-auto
for DS in "${DATASETS[@]}"; do
    echo "--- $DS $(date)"
    python examples/05_tabarena_protocol.py --datasets "$DS" --model "kumo-tabular-s:n_estimators=8,device=cuda" "${HARNESS_ARGS[@]}" \
        --budget-evals 64 --budget-minutes 150 --cv-repeats auto --name "J10_${TAG}_b64_${DS}" || echo "[J10] $DS FAILED"
    ( cd .. && collect_and_push "J10: budget64 $TAG $DS (job ${SLURM_JOB_ID})" )
done
echo "=== J10 $TAG done $(date) ==="
