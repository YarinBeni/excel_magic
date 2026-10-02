#!/bin/bash
#SBATCH --job-name=J3_vllm
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=08:00:00
#SBATCH --output=sbatch/logs/J3_vllm_%j.out
# J3 — the LLM search with an OPEN model, no API: vLLM serves Qwen3-Coder-30B-A3B-Instruct on this GPU (own venv),
# the openai harness drives it; frozen Kumo Tabular-S shares the GPU. Synthetic datasets first, then 3 small TabArena
# datasets (Lite) so the LLM vs heuristic gap is measured on the paper's protocol.
set -uo pipefail
echo "=== J3 vllm search === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh
LLM="${LLM:-Qwen/Qwen3-Coder-30B-A3B-Instruct}"; PORT=8000
start_vllm "${PARSERS:-qwen3_coder hermes}" || { echo "FAILED: vllm could not start"; exit 1; }
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
cd open-tabfm-auto
export OPENAI_BASE_URL=http://localhost:8000/v1 OPENAI_API_KEY=EMPTY
for ds in synth_physics synth_entities breast_cancer; do
    tabfm-auto search --dataset $ds --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=4,device=cuda" \
        --budget-evals 16 --budget-minutes 40 --baselines hgb --run-name "J3_openai_$ds" || echo "[J3] $ds FAILED"
done
python examples/05_tabarena_protocol.py --datasets blood-transfusion-service-center,credit-g,airfoil_self_noise --lite \
    --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=8,device=cuda" --budget-evals 24 --budget-minutes 60 \
    --name J3_tabarena_openai || echo "[J3] tabarena FAILED"
cd ..
collect_and_push "J3: open-LLM (vllm) search results (job ${SLURM_JOB_ID})"
echo "=== J3 done $(date) ==="
