#!/bin/bash
#SBATCH --job-name=J5_llm_sweep
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=08:00:00
#SBATCH --array=0-3
#SBATCH --output=sbatch/logs/J5_llm_sweep_%A_%a.out
# J5 — same search, different open LLMs from Hugging Face (one per array task), each served by vLLM on its own GPU.
# Datasets: the 3 synthetic/bundled ones + 3 small TabArena datasets (Lite), frozen Kumo Tabular-S. Compare with
# J1 (heuristic, no LLM) and J3. Add models by extending the arrays below (tool parser must match vLLM's list).
set -uo pipefail
echo "=== J5 llm sweep task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
LLMS=("Qwen/Qwen3-Coder-30B-A3B-Instruct" "Qwen/Qwen3-32B" "zai-org/GLM-4.5-Air" "openai/gpt-oss-20b")
PARSERS=("hermes" "hermes" "glm45" "openai")
EXTRA=("" "--reasoning-parser qwen3" "" "")
LLM="${LLMS[$SLURM_ARRAY_TASK_ID]}"; PARSER="${PARSERS[$SLURM_ARRAY_TASK_ID]}"; XTRA="${EXTRA[$SLURM_ARRAY_TASK_ID]}"
TAG=$(echo "$LLM" | sed 's#.*/##; s/[^0-9A-Za-z_.-]/_/g')
VLLM_VENV="$HOME/venvs/vllm"
if [ ! -f "$VLLM_VENV/bin/vllm" ]; then
    python3 -m venv "$VLLM_VENV" && "$VLLM_VENV/bin/pip" install -q --upgrade pip && "$VLLM_VENV/bin/pip" install -q vllm || { echo "FAILED vllm install"; exit 1; }
fi
mkdir -p artifacts/vllm
PORT=$((8100 + SLURM_ARRAY_TASK_ID))
"$VLLM_VENV/bin/vllm" serve "$LLM" --port $PORT --enable-auto-tool-choice --tool-call-parser "$PARSER" $XTRA \
    --max-model-len 32768 --gpu-memory-utilization 0.70 > artifacts/vllm/J5_${TAG}_${SLURM_JOB_ID}.log 2>&1 &
VLLM_PID=$!
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
for i in $(seq 1 160); do
    curl -sf http://localhost:$PORT/v1/models >/dev/null 2>&1 && break
    kill -0 $VLLM_PID 2>/dev/null || { echo "FAILED: vllm died for $LLM"; tail -40 artifacts/vllm/J5_${TAG}_${SLURM_JOB_ID}.log; exit 1; }
    sleep 15
done
curl -sf http://localhost:$PORT/v1/models >/dev/null || { echo "FAILED: vllm not up after 40 min"; exit 1; }
echo "[J5] $LLM up on :$PORT"
cd open-tabfm-auto
export OPENAI_BASE_URL=http://localhost:$PORT/v1 OPENAI_API_KEY=EMPTY
for ds in synth_physics synth_entities breast_cancer; do
    tabfm-auto search --dataset $ds --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=4,device=cuda" \
        --budget-evals 16 --budget-minutes 40 --baselines hgb --run-name "J5_${TAG}_$ds" || echo "[J5] $ds FAILED"
done
python examples/05_tabarena_protocol.py --datasets blood-transfusion-service-center,credit-g,airfoil_self_noise --lite \
    --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=8,device=cuda" --budget-evals 24 --budget-minutes 60 \
    --name "J5_${TAG}_tabarena" || echo "[J5] tabarena FAILED"
cd ..
collect_and_push "J5: LLM sweep $TAG (job ${SLURM_JOB_ID})"
echo "=== J5 $TAG done $(date) ==="
