#!/bin/bash
#SBATCH --job-name=J6_cli_agents
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=08:00:00
#SBATCH --array=0-2
#SBATCH --output=sbatch/logs/J6_cli_agents_%A_%a.out
# J6 — free coding-agent CLIs as the harness (one per array task: pi, Qwen Code, aider), all driving the SAME open model
# (Qwen3-Coder-30B-A3B via vLLM on this GPU). Node 22 comes from a conda env created on first use.
set -uo pipefail
echo "=== J6 cli agents task ${SLURM_ARRAY_TASK_ID} === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
AGENTS=("pi" "qwen-code" "aider")
AGENT="${AGENTS[$SLURM_ARRAY_TASK_ID]}"
LLM="${LLM:-Qwen/Qwen3-Coder-30B-A3B-Instruct}"; SERVED="$LLM"
PORT=$((8200 + SLURM_ARRAY_TASK_ID))
source sbatch/_vllm.sh
start_vllm "${PARSERS:-qwen3_coder hermes}" || { echo "FAILED: vllm could not start"; exit 1; }
trap 'kill $VLLM_PID 2>/dev/null || true' EXIT
BASE="http://localhost:$PORT/v1"
if [ "$AGENT" = "pi" ] || [ "$AGENT" = "qwen-code" ]; then
    if [ ! -x "$HOME/miniconda3/envs/node22/bin/node" ]; then
        conda create -y -q -n node22 -c conda-forge "nodejs>=22" || { echo "FAILED conda node22"; exit 1; }
    fi
    export PATH="$HOME/miniconda3/envs/node22/bin:$PATH"
    node --version
fi
case "$AGENT" in
  pi)
    command -v pi >/dev/null || npm i -g --ignore-scripts @earendil-works/pi-coding-agent || { echo "FAILED npm pi"; exit 1; }
    mkdir -p ~/.pi/agent
    cat > ~/.pi/agent/models.json <<JSON
{"providers":{"vllm":{"baseUrl":"$BASE","api":"openai-completions","apiKey":"dummy","models":[{"id":"$SERVED"}]}}}
JSON
    ;;
  qwen-code)
    command -v qwen >/dev/null || npm i -g @qwen-code/qwen-code@latest || { echo "FAILED npm qwen-code"; exit 1; }
    export QWEN_CODE_SUPPRESS_YOLO_WARNING=1 QWEN_CODE_UNATTENDED_RETRY=1
    ;;
  aider)
    command -v aider >/dev/null || python -m pip install -q aider-chat || { echo "FAILED pip aider"; exit 1; }
    ;;
esac
cd open-tabfm-auto
export OPENAI_BASE_URL="$BASE" OPENAI_API_BASE="$BASE" OPENAI_API_KEY=dummy OPENAI_MODEL="$SERVED"
if [ "$AGENT" = "aider" ]; then
    # aider is single-turn: wrap it in a small loop (write -> eval -> feed the score back) as the agent command
    cat > /tmp/aider_loop_$SLURM_JOB_ID.sh <<'SH'
#!/bin/bash
PROMPT="$1"; MODEL="$2"
for i in 1 2 3 4 5 6 7 8; do
  aider --message-file "$PROMPT" --yes-always --no-git --no-auto-commits --no-show-model-warnings --no-check-update --no-analytics \
        --model "openai/$MODEL" pipeline.py || true
  tabfm-eval | tee -a NOTES.md | tail -4
  grep -q "BUDGET EXHAUSTED" NOTES.md && break
  { echo; echo "## Evaluation $i result (improve on it; keep what helped, revert what hurt):"; tail -6 NOTES.md; } >> "$PROMPT"
done
SH
    chmod +x /tmp/aider_loop_$SLURM_JOB_ID.sh
    CMD="/tmp/aider_loop_$SLURM_JOB_ID.sh {prompt_file} {model}"
else
    CMD="$AGENT"
fi
for ds in synth_physics synth_entities breast_cancer; do
    tabfm-auto search --dataset $ds --harness cli --agent-cmd "$CMD" --llm "$SERVED" --llm-base-url "$BASE" \
        --model "kumo-tabular-s:n_estimators=4,device=cuda" --budget-evals 16 --budget-minutes 40 --baselines hgb \
        --run-name "J6_${AGENT}_$ds" || echo "[J6] $AGENT $ds FAILED"
done
cd ..
collect_and_push "J6: cli agent $AGENT (job ${SLURM_JOB_ID})"
echo "=== J6 $AGENT done $(date) ==="
