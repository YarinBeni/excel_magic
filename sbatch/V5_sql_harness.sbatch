#!/bin/bash
#SBATCH --job-name=V5_sql_harness
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:2
#SBATCH --cpus-per-task=24
#SBATCH --mem=200G
#SBATCH --time=16:00:00
#SBATCH --output=sbatch/logs/V5_sql_harness_%j.out
# V5 (phases P2 + P4): SQL harness ablation A, B, C, D, F, SC on BIRD Arcwise-Plat (generator Qwen3-Coder-30B on GPU 0,
# judge GLM-4.5-Air on GPU 1), then the guarded auto-research loop over the harness knobs. Needs SIGNALS=<P0 signals.csv>.
set -uo pipefail
echo "=== V5 sql harness === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
export BIRD_ROOT="$REPO/artifacts/bird"
SIG="${SIGNALS:?set SIGNALS to the P0 signals.csv}"
source sbatch/_vllm.sh
CUDA_VISIBLE_DEVICES=0 LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8240 VLLM_GPU_FRAC=0.80 VLLM_MAX_LEN=32768 start_vllm "qwen3_coder hermes" \
  || { echo "FAILED vllm gen"; collect_and_push "V5: generator vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
GEN_PID=$VLLM_PID
CUDA_VISIBLE_DEVICES=1 LLM=zai-org/GLM-4.5-Air-FP8 PORT=8241 VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=32768 start_vllm "glm45 hermes" || echo "WARN judge vLLM failed: F skipped"
JUDGE_PID=$VLLM_PID
trap 'kill $GEN_PID $JUDGE_PID 2>/dev/null || true' EXIT
export OPENAI_BASE_URL=http://localhost:8240/v1 OPENAI_API_KEY=EMPTY JUDGE_BASE_URL=http://localhost:8241/v1
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V5_sql_harness"; mkdir -p "$RUN"
if [ -z "${SKIP_P2:-}" ]; then
  echo "--- P2 ablation $(date)"
  CUDA_VISIBLE_DEVICES=0 python vfd-agent/experiments/p2_sql_harness.py --run "$RUN" --signals "$SIG" --cache artifacts/vfd_data \
      --configs ${CONFIGS:-A,B,C,D,F,SC} ${LIMIT:+--limit $LIMIT} || echo "FAILED p2"
  ( cd "$REPO" && collect_and_push "V5: SQL harness ablation (job ${SLURM_JOB_ID})" )
fi
if [ -z "${SKIP_P4:-}" ]; then
  echo "--- P4 auto-research $(date)"
  RUN4="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V5_autoresearch_sql"; mkdir -p "$RUN4"
  CUDA_VISIBLE_DEVICES=0 RESEARCHER_BASE_URL=http://localhost:8241/v1 python vfd-agent/experiments/p4_autoresearch_sql.py \
      --run "$RUN4" --signals "$SIG" --cache artifacts/vfd_data --budget ${BUDGET:-20} || echo "FAILED p4"
fi
collect_and_push "V5: SQL harness ablation + guarded auto-research (job ${SLURM_JOB_ID})"
echo "=== V5 done $(date) ==="
