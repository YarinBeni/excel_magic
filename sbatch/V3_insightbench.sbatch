#!/bin/bash
#SBATCH --job-name=V3_insightbench
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=16:00:00
#SBATCH --output=sbatch/logs/V3_insightbench_%j.out
# V3 (phase P3, claim C3): InsightBench, analyst agent configs D (LLM+SQL), E (+DEEP tools), F (+GLiClass text labels and
# verified ledger). Generator Qwen3-Coder-30B; judge GLM-4.5-Air (other family). LIMIT=n for a smoke run.
set -uo pipefail
echo "=== V3 insightbench === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
export INSIGHTBENCH="$REPO/artifacts/insight-bench"
[ -d "$INSIGHTBENCH/data/notebooks" ] || git clone -q --depth 1 https://github.com/ServiceNow/insight-bench.git "$INSIGHTBENCH"
python -c "import duckdb" 2>/dev/null || pip install -q duckdb 2>&1 | tail -1
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V3_insightbench_${LIMIT:-all}"; mkdir -p "$RUN"
P="python vfd-agent/experiments/p3_insightbench.py --run $RUN ${LIMIT:+--limit $LIMIT}"
source sbatch/_vllm.sh
echo "--- analyze $(date)"
LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8220 VLLM_GPU_FRAC=0.60 VLLM_MAX_LEN=65536 start_vllm "qwen3_coder hermes" \
  || { echo "FAILED vllm"; collect_and_push "V3: vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
OPENAI_BASE_URL=http://localhost:8220/v1 OPENAI_API_KEY=EMPTY $P analyze --configs ${CONFIGS:-D,E,F} --workers 8 || echo "FAILED analyze"
kill $VLLM_PID 2>/dev/null; sleep 20
echo "--- judge $(date)"
LLM=zai-org/GLM-4.5-Air-FP8 PORT=8221 VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=32768 start_vllm "glm45 hermes" \
  && { OPENAI_BASE_URL=http://localhost:8221/v1 OPENAI_API_KEY=EMPTY $P judge --model zai-org/GLM-4.5-Air-FP8 --workers 32 || echo "FAILED judge"; } \
  || echo "FAILED vllm judge"
kill $VLLM_PID 2>/dev/null
$P report || echo "FAILED report"
collect_and_push "V3: InsightBench analyst ablation D/E/F (job ${SLURM_JOB_ID})"
echo "=== V3 done $(date) ==="
