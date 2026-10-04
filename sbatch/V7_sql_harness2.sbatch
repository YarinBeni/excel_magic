#!/bin/bash
#SBATCH --job-name=V7_sql_h2
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=24
#SBATCH --mem=200G
#SBATCH --time=12:00:00
#SBATCH --output=sbatch/logs/V7_sql_h2_%j.out
# V7 (Q1/Q2): harness v2 ladder on BIRD Arcwise-Plat-Full. One model per job (MODEL, PARSERS, TAG); the corrected
# column descriptions come from the Arcwise repo (uiuc-kang-lab/text_to_sql_benchmarks), downloaded once.
#   MODEL=Qwen/Qwen3-Coder-30B-A3B-Instruct CONFIGS=A,K1,K2,K3,G,GD SPLIT=dev sbatch sbatch/V7_sql_harness2.sbatch
set -uo pipefail
echo "=== V7 sql harness v2 === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
export BIRD_ROOT="$REPO/artifacts/bird"
DESC="$REPO/artifacts/bird_desc"
if [ ! -f "$DESC/.ok" ]; then
  rm -rf /tmp/ttsb_$USER; git clone -q --depth 1 https://github.com/uiuc-kang-lab/text_to_sql_benchmarks /tmp/ttsb_$USER
  find /tmp/ttsb_$USER -type d -name database_description | while read d; do
    db=$(basename "$(dirname "$d")"); mkdir -p "$DESC/$db"; cp -r "$d" "$DESC/$db/"; done
  ls "$DESC" && touch "$DESC/.ok"
fi
echo "descriptions for: $(ls $DESC | tr '\n' ' ')"
MODEL="${MODEL:-Qwen/Qwen3-Coder-30B-A3B-Instruct}"
source sbatch/_vllm.sh
LLM="$MODEL" PORT=8260 VLLM_GPU_FRAC="${VLLM_GPU_FRAC:-0.85}" VLLM_MAX_LEN="${VLLM_MAX_LEN:-32768}" start_vllm "${PARSERS:-qwen3_coder hermes}" "${VLLM_EXTRA:-}" \
  || { echo "FAILED vllm"; collect_and_push "V7: vLLM failed for $MODEL (job ${SLURM_JOB_ID})"; exit 1; }
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
export OPENAI_BASE_URL=http://localhost:8260/v1 OPENAI_API_KEY=EMPTY
RUN="${RUN:-$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V7_sql_harness2}"; mkdir -p "$RUN"
python vfd-agent/experiments/p5_sql_harness2.py --run "$RUN" --cache artifacts/vfd_data --desc-root "$DESC" \
    --split "${SPLIT:-dev}" --configs "${CONFIGS:-A,K1,K2,K3,G,GD}" --model "$MODEL" ${TAG:+--tag $TAG} || echo "FAILED harness2"
collect_and_push "V7: harness v2 ladder, $MODEL, ${SPLIT:-dev} (job ${SLURM_JOB_ID})"
echo "=== V7 done $(date) ==="
