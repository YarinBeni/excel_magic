#!/bin/bash
#SBATCH --job-name=V4_rel_capability
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=200G
#SBATCH --time=20:00:00
#SBATCH --output=sbatch/logs/V4_rel_capability_%j.out
# V4 (phase P3, claims C4/C7): prediction questions over RelBench databases answered by an agent (Qwen3-Coder-30B):
# D = LLM + SQL, E = + Kumo Tabular-L on LLM-built features, R = + Kumo Relational graph-layer probe, ER = both.
set -uo pipefail
echo "=== V4 relbench capability === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
python -c "import duckdb" 2>/dev/null || pip install -q duckdb 2>&1 | tail -1
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V4_relbench_capability"; mkdir -p "$RUN"
source sbatch/_vllm.sh
LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8230 VLLM_GPU_FRAC=0.55 VLLM_MAX_LEN=65536 start_vllm "qwen3_coder hermes" \
  || { echo "FAILED vllm"; collect_and_push "V4: vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
OPENAI_BASE_URL=http://localhost:8230/v1 OPENAI_API_KEY=EMPTY python vfd-agent/experiments/p3_relbench_capability.py \
    --run "$RUN" ${TASKS:+--tasks $TASKS} --configs ${CONFIGS:-D,E,R,ER} --repeats ${REPEATS:-2} || echo "FAILED capability"
collect_and_push "V4: RelBench prediction questions, agent D/E/R/ER (job ${SLURM_JOB_ID})"
echo "=== V4 done $(date) ==="
