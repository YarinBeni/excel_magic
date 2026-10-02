#!/usr/bin/env bash
# One TabArena dataset per SLURM array task on a GPU node. Adjust partition/account; needs ANTHROPIC_API_KEY
# (or a logged-in `claude` CLI) and api.openml.org / huggingface.co reachable from the node.
#SBATCH --job-name=tabfm-auto
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=08:00:00
#SBATCH --array=0-50
#SBATCH --output=logs/tabarena_%a.out
set -euo pipefail
cd "$(dirname "$0")/.."
source .venv/bin/activate
export OPENML_CACHE_DIR=${OPENML_CACHE_DIR:-$PWD/data_cache/openml}
export TABFM_RUNS_ROOT=${TABFM_RUNS_ROOT:-$PWD/runs}
DATASET=$(python -c "from tabfm_auto.benchmarks.tabarena import list_datasets; print(list_datasets()[${SLURM_ARRAY_TASK_ID}].name)")
MODEL=${MODEL:-kumo-tabular-s:n_estimators=8,device=cuda}
python examples/05_tabarena_protocol.py --datasets "$DATASET" --model "$MODEL" --llm "${LLM:-opus}" \
  --budget-evals "${BUDGET_EVALS:-96}" --budget-minutes "${BUDGET_MINUTES:-360}" --name "tabarena_${DATASET}"
