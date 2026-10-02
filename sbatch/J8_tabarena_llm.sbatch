#!/bin/bash
#SBATCH --job-name=J8_tabarena_llm
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=14:00:00
#SBATCH --array=0-1
#SBATCH --output=sbatch/logs/J8_tabarena_llm_%A_%a.out
# J8 — the paper's TabArena protocol (17 datasets <= 2500 rows, search on r0f0 train, score all official splits) with the
# two best free setups from J5/J6, one per array task, frozen Kumo Tabular-S (n_estimators=8):
#   task 0: pi CLI agent + Qwen3-Coder-30B-A3B (vLLM)      task 1: our tool loop + GLM-4.5-Air-FP8 (vLLM)
# One dataset at a time, results pushed after each dataset. Compare with J2 (no-LLM heuristic) and the paper's Table 9.
set -uo pipefail
echo "=== J8 tabarena llm task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh
PORT=$((8300 + SLURM_ARRAY_TASK_ID))
if [ "$SLURM_ARRAY_TASK_ID" = "0" ]; then
    LLM="Qwen/Qwen3-Coder-30B-A3B-Instruct"; TAG="pi_qwen3coder"
    start_vllm "qwen3_coder hermes" || { echo "FAILED: vllm could not start"; exit 1; }
else
    LLM="zai-org/GLM-4.5-Air-FP8"; TAG="openai_glm45air"
    export VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=65536   # FP8 weights leave ~18 GiB for KV on one H200: 128k needs 23 GiB
    start_vllm "glm45 hermes" || { echo "FAILED: vllm could not start"; exit 1; }
fi
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
BASE="http://localhost:$PORT/v1"
export OPENAI_BASE_URL="$BASE" OPENAI_API_BASE="$BASE" OPENAI_API_KEY=dummy OPENAI_MODEL="$LLM"
if [ "$SLURM_ARRAY_TASK_ID" = "0" ]; then
    export PATH="$HOME/miniconda3/envs/node22/bin:$PATH"
    command -v pi >/dev/null || npm i -g --ignore-scripts @earendil-works/pi-coding-agent || { echo "FAILED npm pi"; exit 1; }
    mkdir -p ~/.pi/agent
    cat > ~/.pi/agent/models.json <<JSON
{"providers":{"vllm":{"baseUrl":"$BASE","api":"openai-completions","apiKey":"dummy","models":[{"id":"$LLM"}]}}}
JSON
    HARNESS_ARGS=(--harness cli --agent-cmd pi)
else
    HARNESS_ARGS=(--harness openai)
fi
cd open-tabfm-auto
mapfile -t DATASETS < <(python -c "from tabfm_auto.benchmarks.tabarena import list_datasets; print('\n'.join(d.name for d in list_datasets(max_instances=2500, max_features=200)))")
echo "${#DATASETS[@]} datasets: ${DATASETS[*]}"
for DS in "${DATASETS[@]}"; do
    echo "--- $DS $(date)"
    python examples/05_tabarena_protocol.py --datasets "$DS" --model "kumo-tabular-s:n_estimators=8,device=cuda" \
        "${HARNESS_ARGS[@]}" --llm "$LLM" --llm-base-url "$BASE" --budget-evals "${BUDGET_EVALS:-16}" \
        --budget-minutes "${BUDGET_MINUTES:-40}" --name "J8_${TAG}_${DS}" || echo "[J8] $DS FAILED"
    ( cd .. && collect_and_push "J8: tabarena $TAG $DS (job ${SLURM_JOB_ID})" )
done
echo "=== J8 $TAG done $(date) ==="
