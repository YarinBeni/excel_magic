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
VLLM_VENV="$HOME/venvs/vllm"
LLM="${LLM:-Qwen/Qwen3-Coder-30B-A3B-Instruct}"
if [ ! -f "$VLLM_VENV/bin/vllm" ]; then
    echo "[J3] installing vllm into $VLLM_VENV (one-time, ~10 min)"
    python3 -m venv "$VLLM_VENV" && "$VLLM_VENV/bin/pip" install -q --upgrade pip && "$VLLM_VENV/bin/pip" install -q vllm || { echo "FAILED vllm install"; exit 1; }
fi
mkdir -p artifacts/vllm
"$VLLM_VENV/bin/vllm" serve "$LLM" --port 8000 --enable-auto-tool-choice --tool-call-parser hermes \
    --max-model-len 32768 --gpu-memory-utilization 0.70 > artifacts/vllm/server_${SLURM_JOB_ID}.log 2>&1 &
VLLM_PID=$!
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
for i in $(seq 1 120); do
    curl -sf http://localhost:8000/v1/models >/dev/null 2>&1 && break
    kill -0 $VLLM_PID 2>/dev/null || { echo "FAILED: vllm died"; tail -40 artifacts/vllm/server_${SLURM_JOB_ID}.log; exit 1; }
    sleep 15
done
curl -sf http://localhost:8000/v1/models | head -c 400 || { echo "FAILED: vllm not up after 30 min"; exit 1; }
echo; echo "[J3] vllm up"
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
