#!/bin/bash
#SBATCH --job-name=V10_gli_cmp
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=6:00:00
#SBATCH --output=sbatch/logs/V10_gli_cmp_%j.out
# V10 (Q5b): zero-shot GLiClass vs GLiNER2 / GLiNER2.5 / GLiNER2.5-Decide (open weights) as SQL checker and column linker.
set -uo pipefail
echo "=== V10 gli compare === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
python -c "import gliner2, peft" 2>/dev/null || pip install -q --no-deps "gliner2>=2.0" peft 2>&1 | tail -2
python -c "import transformers, gliner2; print('transformers', transformers.__version__, 'gliner2', gliner2.__version__)"
export BIRD_ROOT="$REPO/artifacts/bird"
P0=$(ls -td artifacts/runs/*V1_verifier_study_all | head -1)
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V10_gli_compare"; mkdir -p "$RUN"
python vfd-agent/experiments/p8_gli_compare.py --run "$RUN" --p0 "$P0" --cache artifacts/vfd_data \
    --desc-root "$REPO/artifacts/bird_desc" ${MODELS:+--models $MODELS} || echo "FAILED compare"
collect_and_push "V10: zero-shot GLi-family comparison (job ${SLURM_JOB_ID})"
echo "=== V10 done $(date) ==="
