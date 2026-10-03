#!/bin/bash
#SBATCH --job-name=J15b_selection_glm
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=96G
#SBATCH --time=06:00:00
#SBATCH --output=sbatch/logs/J15b_selection_glm_%j.out
# J15 — fix the noisy final pick without re-running any agent: re-score the finished TabArena searches (J2 heuristic,
# J8 pi + Qwen3-Coder, J8 GLM loop) on all official splits under conservative selection rules:
#   p0 | best (the paper's rule) | gated1 / gated2 (paired per-fold improvement over P0 > 1 / 2 standard errors, else P0)
#   | ens3 (average the 3 best-CV candidates' predictions).
set -uo pipefail
echo "=== J15b selection rules === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd open-tabfm-auto
for pat in J8_openai_glm45air_; do
    runs=$(ls -d ../artifacts/runs/*${pat}* 2>/dev/null | grep -v "_smoke" || true)
    [ -n "$runs" ] || { echo "no runs for $pat"; continue; }
    python scripts/rescore_selection.py --runs $runs --rules best,gated1,gated2,ens3 || echo "[J15] $pat FAILED"
    ( cd .. && collect_and_push "J15b: selection rules ${pat%_} (job ${SLURM_JOB_ID})" )
done
echo "=== J15b done $(date) ==="
