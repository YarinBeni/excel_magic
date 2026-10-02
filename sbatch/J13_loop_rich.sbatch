#!/bin/bash
#SBATCH --job-name=J13_loop_rich
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=04:00:00
#SBATCH --output=sbatch/logs/J13_loop_rich_%j.out
# J13 — is the harness gap affordances or agency? The minimal tool loop with the CLI agents' context (file listing,
# pipeline.py, data summary) in the first message (TABFM_LOOP_RICH=1), same Qwen3-Coder-30B, same 6 tasks as J5.
set -uo pipefail
echo "=== J13 loop rich === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh
LLM="Qwen/Qwen3-Coder-30B-A3B-Instruct"; PORT=8600; BASE="http://localhost:$PORT/v1"
start_vllm "qwen3_coder hermes" || { echo "FAILED: vllm"; exit 1; }; trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
export OPENAI_BASE_URL="$BASE" OPENAI_API_KEY=EMPTY TABFM_LOOP_RICH=1
cd open-tabfm-auto
for ds in synth_physics synth_entities breast_cancer; do
    tabfm-auto search --dataset $ds --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=4,device=cuda" \
        --budget-evals 16 --budget-minutes 40 --baselines hgb --run-name "J13_rich_qwen3coder_$ds" || echo "[J13] $ds FAILED"
done
python examples/05_tabarena_protocol.py --datasets blood-transfusion-service-center,credit-g,airfoil_self_noise --lite \
    --harness openai --llm "$LLM" --model "kumo-tabular-s:n_estimators=8,device=cuda" --budget-evals 24 --budget-minutes 60 \
    --name "J13_rich_qwen3coder_tabarena" || echo "[J13] tabarena FAILED"
cd ..
collect_and_push "J13: rich tool loop Qwen3-Coder (job ${SLURM_JOB_ID})"
echo "=== J13 done $(date) ==="
