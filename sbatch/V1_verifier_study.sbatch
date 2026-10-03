#!/bin/bash
#SBATCH --job-name=V1_verifier_study
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=10:00:00
#SBATCH --output=sbatch/logs/V1_verifier_study_%j.out
# V1 (phase P0 of the three-model harness study): on BIRD Arcwise-Plat (498 fully corrected questions), how well do cheap
# signals tell a correct SQL candidate from a wrong one, and find the columns a question needs?
#   A  Qwen3-Coder-30B (vLLM) writes 1 greedy + 8 sampled SQL per question; each is executed and labelled vs gold
#   B  GLM-4.5-Air (vLLM, other family, thinking off) judges every candidate: P(yes)
#   C  GLiClass-large-v3, NLI cross-encoder; schema linking: GLiClass, GLiNER bi-encoder, bge reranker, bge-small, lexical
#   D  report: AUROC, best-of-N, ECE raw / isotonic, stacking, cascade, linking recall@k, latency
set -uo pipefail
echo "=== V1 verifier study === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
export BIRD_ROOT="$REPO/artifacts/bird" HF_HUB_DOWNLOAD_TIMEOUT=60
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V1_verifier_study_${LIMIT:-all}"; mkdir -p "$RUN"
P="python vfd-agent/experiments/p0_verifier_study.py --run $RUN --cache artifacts/vfd_data ${LIMIT:+--limit $LIMIT}"
source sbatch/_vllm.sh
echo "--- A generate $(date)"
LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8200 VLLM_GPU_FRAC=0.85 VLLM_MAX_LEN=32768 start_vllm "qwen3_coder hermes" \
    || { echo "FAILED vllm generator"; collect_and_push "V1: generator vLLM failed (job ${SLURM_JOB_ID})"; exit 1; }
OPENAI_BASE_URL=http://localhost:8200/v1 OPENAI_API_KEY=EMPTY $P generate --model Qwen/Qwen3-Coder-30B-A3B-Instruct --n 8 --workers 48 \
    || echo "FAILED generate"
kill $VLLM_PID 2>/dev/null; sleep 20
[ -s "$RUN/candidates.jsonl" ] || { echo "FAILED no candidates"; collect_and_push "V1: no candidates (job ${SLURM_JOB_ID})"; exit 1; }
echo "--- B judge $(date)"
LLM=zai-org/GLM-4.5-Air-FP8 PORT=8201 VLLM_GPU_FRAC=0.90 VLLM_MAX_LEN=32768 start_vllm "glm45 hermes" \
    && { OPENAI_BASE_URL=http://localhost:8201/v1 OPENAI_API_KEY=EMPTY $P judge --model zai-org/GLM-4.5-Air-FP8 --workers 48 || echo "FAILED judge"; } \
    || echo "FAILED vllm judge"
kill $VLLM_PID 2>/dev/null; sleep 20
echo "--- C encoders $(date)"
$P encoders || echo "FAILED encoders"
echo "--- C link $(date)"
$P link || echo "FAILED link"
echo "--- D report $(date)"
$P report || echo "FAILED report"
collect_and_push "V1: verifier study on BIRD Arcwise-Plat (job ${SLURM_JOB_ID})"
echo "=== V1 done $(date) ==="
