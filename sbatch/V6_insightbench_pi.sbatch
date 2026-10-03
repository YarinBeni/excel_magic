#!/bin/bash
#SBATCH --job-name=V6_insight_pi
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=20:00:00
#SBATCH --output=sbatch/logs/V6_insight_pi_%j.out
# V6 (phase P3, harness swap): InsightBench with the pi agent (the company's harness) using the same three-model tools as
# V3, through the `vfd` shell command and one tool server per config (D, E, F). Generator Qwen3-Coder-30B; judge GLM-4.5-Air.
set -uo pipefail
echo "=== V6 insightbench pi === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
export INSIGHTBENCH="$REPO/artifacts/insight-bench"
[ -d "$INSIGHTBENCH/data/notebooks" ] || git clone -q --depth 1 https://github.com/ServiceNow/insight-bench.git "$INSIGHTBENCH"
python -c "import duckdb" 2>/dev/null || pip install -q duckdb 2>&1 | tail -1
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V6_insightbench_pi_${LIMIT:-all}"; mkdir -p "$RUN"
source sbatch/_vllm.sh; source sbatch/_agents.sh
LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8250 VLLM_GPU_FRAC=0.55 VLLM_MAX_LEN=65536 start_vllm "qwen3_coder hermes" \
  || { echo "FAILED vllm"; collect_and_push "V6: vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
GEN=$VLLM_PID
setup_pi "http://localhost:8250/v1" "Qwen/Qwen3-Coder-30B-A3B-Instruct" || { echo "FAILED pi"; exit 1; }
export PYTHONPATH="$REPO/vfd-agent:${PYTHONPATH:-}"
mkdir -p "$RUN/bin"; printf '#!/bin/bash\nexec python -c "from vfd.toolserver import cli; cli()" "$@"\n' > "$RUN/bin/vfd"; chmod +x "$RUN/bin/vfd"
export PATH="$RUN/bin:$PATH"
SP=()
for c in D E F; do port=$((8765 + $(printf '%d' "'$c") - 68)); python -m vfd.toolserver --config $c --port $port > "$RUN/toolserver_$c.log" 2>&1 & SP+=($!); done
trap 'kill $GEN ${SP[@]} 2>/dev/null || true' EXIT
sleep 90   # models load
python vfd-agent/experiments/p3_insightbench_pi.py --run "$RUN" --configs ${CONFIGS:-D,E,F} --ports D=8765,E=8766,F=8767 \
    ${LIMIT:+--limit $LIMIT} || echo "FAILED pi runs"
kill $GEN ${SP[@]} 2>/dev/null; sleep 20
LLM=zai-org/GLM-4.5-Air-FP8 PORT=8251 VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=32768 start_vllm "glm45 hermes" \
  && { OPENAI_BASE_URL=http://localhost:8251/v1 OPENAI_API_KEY=EMPTY python vfd-agent/experiments/p3_insightbench.py judge \
       --run "$RUN" --model zai-org/GLM-4.5-Air-FP8 --workers 32 ${LIMIT:+--limit $LIMIT} || echo "FAILED judge"; } || echo "FAILED judge vllm"
kill $VLLM_PID 2>/dev/null
python vfd-agent/experiments/p3_insightbench.py report --run "$RUN" || echo "FAILED report"
rm -rf "$RUN"/ws_pi*/*/data.csv
collect_and_push "V6: InsightBench via pi with the three-model tools (job ${SLURM_JOB_ID})"
echo "=== V6 done $(date) ==="
