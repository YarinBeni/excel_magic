#!/bin/bash
#SBATCH --job-name=V8_rel_hypo
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=200G
#SBATCH --time=20:00:00
#SBATCH --output=sbatch/logs/V8_rel_hypo_%j.out
# V8 (Q3): LLM hypotheses tested on top of the frozen Kumo Relational FM (vfd.hypo).

set -uo pipefail
echo "=== V8 relbench hypotheses === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
python -c "import duckdb" 2>/dev/null || pip install -q duckdb 2>&1 | tail -1
RUN="${RUN:-$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V8_relbench_hypo}"; mkdir -p "$RUN"
source sbatch/_vllm.sh
MODEL="${MODEL:-Qwen/Qwen3-Coder-30B-A3B-Instruct}"
LLM="$MODEL" PORT=8230 VLLM_GPU_FRAC=0.55 VLLM_MAX_LEN=65536 start_vllm "${PARSERS:-qwen3_coder hermes}" \
  || { echo "FAILED vllm"; collect_and_push "V8: vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
OPENAI_BASE_URL=http://localhost:8230/v1 OPENAI_API_KEY=EMPTY python vfd-agent/experiments/p6_relbench_hypo.py \
    --run "$RUN" ${TASKS:+--tasks $TASKS} --repeats ${REPEATS:-2} --budget ${BUDGET:-10} --model "$MODEL" ${TAG:+--tag $TAG} || echo "FAILED hypo"
collect_and_push "V8: RelBench LLM hypotheses on the frozen FM (job ${SLURM_JOB_ID})"
echo "=== V8 done $(date) ==="
