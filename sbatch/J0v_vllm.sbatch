#!/bin/bash
#SBATCH --job-name=J0v_vllm
#SBATCH --partition=voltagepark
#SBATCH --cpus-per-task=16
#SBATCH --mem=64G
#SBATCH --time=02:00:00
#SBATCH --output=sbatch/logs/J0v_vllm_%j.out
# J0v — install vLLM once (CPU job, under a lock) so the GPU jobs J3/J5/J6 never race on the venv.
set -uo pipefail
echo "=== J0v vllm install === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
source sbatch/_vllm.sh
ensure_vllm || { echo "FAILED"; exit 1; }
vllm_env
"$VLLM_VENV/bin/python" -c "import vllm; print('vllm', vllm.__version__)" | tee reports/J0v_vllm_version.txt
"$VLLM_VENV/bin/python" -c "import torch; print('vllm venv torch', torch.__version__, torch.cuda.is_available())" >> reports/J0v_vllm_version.txt 2>&1 || true
commit_push_sandbox "J0v: vllm installed (job ${SLURM_JOB_ID})" reports
echo "=== J0v done $(date) ==="
