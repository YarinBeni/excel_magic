#!/bin/bash
#SBATCH --job-name=AGENT_runner
#SBATCH --partition=voltagepark
#SBATCH --cpus-per-task=2
#SBATCH --mem=4G
#SBATCH --time=1-00:00:00
#SBATCH --output=sbatch/logs/AGENT_runner_%j.out
# The git-driven command runner (agent/runner.sh) as a ONE-DAY, 2-CPU job. Started per session with
#   sbatch sbatch/AGENT_runner.sbatch        (from ~/projects/excel_magic)
# stopped by pushing an empty agent/STOP or `scancel`. See agent/README.md and docs/CLUSTER_HANDOFF.md.
set -uo pipefail
echo "=== AGENT runner === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic
mkdir -p sbatch/logs
exec bash agent/runner.sh
