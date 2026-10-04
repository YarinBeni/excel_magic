#!/bin/bash
#SBATCH --job-name=V9_gli_ft
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=6:00:00
#SBATCH --output=sbatch/logs/V9_gli_ft_%j.out
# V9 (Q5): fine-tune GLiClass on the BIRD DEV questions only (checker + column linker), test on HELDOUT. LLM unchanged.
set -uo pipefail
echo "=== V9 gliclass fine-tune === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
python -c "import accelerate" 2>/dev/null || pip install -q --no-deps "accelerate>=0.26" 2>&1 | tail -1
export BIRD_ROOT="$REPO/artifacts/bird"
P0=$(ls -td artifacts/runs/*V1_verifier_study_all | head -1)
RUN="${RUN:-$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V9_gliclass_ft}"; mkdir -p "$RUN"
python vfd-agent/experiments/p7_gliclass_finetune.py --run "$RUN" --p0 "$P0" --cache artifacts/vfd_data \
    --desc-root "$REPO/artifacts/bird_desc" --models "$REPO/artifacts/gliclass_ft" --tasks "${TASKS:-checker,linker}" \
    --epochs "${EPOCHS:-3}" || echo "FAILED gliclass ft"
collect_and_push "V9: GLiClass fine-tuned on BIRD dev, tested on held-out (job ${SLURM_JOB_ID})"
echo "=== V9 done $(date) ==="
