#!/bin/bash
#SBATCH --job-name=J0_setup
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=03:00:00
#SBATCH --output=sbatch/logs/J0_setup_%j.out
# J0 — environment + weights + tests. THE gate: nothing else runs until this exits 0.
#   * venv ~/venvs/tabfm on top of the `thesis` conda python (system-site-packages: inherits torch)
#   * pip: open-tabfm-auto[tabpfn,tabicl,baselines,dev], NVIDIA sdm (Kumo Tabular / Relational), EXAONE, torch_geometric,
#          openml, relbench, frozen-embeddings-retrieval (no-deps)
#   * weights: tabfm-models download (Hugging Face reachable from compute nodes) + OpenRFM from GitHub
#   * pytest of both packages
set -uo pipefail
echo "=== J0 setup === job ${SLURM_JOB_ID} $(date)"
source ~/miniconda3/etc/profile.d/conda.sh && conda activate thesis
cd ~/projects/excel_magic
source sbatch/_sandbox.sh
sync_sandbox || { echo "FAILED sync_sandbox"; exit 1; }
enable_log_push
mkdir -p sbatch/logs artifacts/weights reports
VENV="$HOME/venvs/tabfm"
if [ ! -f "$VENV/bin/activate" ]; then
    python -m venv --system-site-packages "$VENV" || { echo "FAILED venv"; exit 1; }
fi
source "$VENV/bin/activate"
python -m pip install -q --upgrade pip
python -c "import torch; print('torch', torch.__version__, 'cuda', torch.cuda.is_available())"
TORCH_OK=$(python -c "import torch; from packaging.version import Version as V; print(int(V(torch.__version__.split('+')[0]) >= V('2.7')))" 2>/dev/null || echo 0)
if [ "$TORCH_OK" != "1" ]; then
    echo "[J0] torch < 2.7 in thesis env; installing a newer torch into the venv for NVIDIA sdm"
    python -m pip install -q "torch>=2.7" || { echo "FAILED torch upgrade"; exit 1; }
fi
python -m pip install -q -e "open-tabfm-auto[tabpfn,tabicl,baselines,dev]" || { echo "FAILED pip open-tabfm-auto"; exit 1; }
python -m pip install -q "sdm @ git+https://github.com/NVIDIA/structured-data-models.git" || echo "[J0] WARN: sdm install failed (Kumo backbones unavailable)"
python -m pip install -q "exaonetabular @ git+https://github.com/LGAI-Research/EXAONE-Tabular.git" || echo "[J0] WARN: exaone install failed"
python -m pip install -q torch_geometric openml relbench tabulate || { echo "FAILED pip extras"; exit 1; }
python -m pip install -q --no-deps -e frozen-embeddings-retrieval || { echo "FAILED pip fer"; exit 1; }
export TABFM_WEIGHTS_DIR="$PWD/artifacts/weights"
export HF_HUB_DOWNLOAD_TIMEOUT=60
cd open-tabfm-auto
tabfm-models download tabpfn tabicl kumo-tabular-s kumo-relational tabpfn-2.5 exaone 2>&1 | tail -20
mkdir -p "$TABFM_WEIGHTS_DIR/openrfm"
[ -s "$TABFM_WEIGHTS_DIR/openrfm/openrfm-pretrain-ctx96.pt" ] || curl -sS -L --fail -o "$TABFM_WEIGHTS_DIR/openrfm/openrfm-pretrain-ctx96.pt" \
    https://github.com/T-Lab/OpenRFM/releases/download/v0.1.0/openrfm-pretrain-ctx96.pt || echo "[J0] WARN: OpenRFM download failed"
tabfm-models status | tee ../reports/J0_models_status.md
python -m pytest -q 2>&1 | tail -3 | tee ../reports/J0_tests_open-tabfm-auto.txt
cd ../frozen-embeddings-retrieval
ln -sfn "$TABFM_WEIGHTS_DIR/openrfm" weights/openrfm; ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
python -m pytest -q 2>&1 | tail -3 | tee ../reports/J0_tests_fer.txt
cd ..
{ echo "# J0 environment ($(date -u))"; python -m pip list 2>/dev/null | grep -iE "^(torch|tabpfn|tabicl|sdm|exaone|torch-geometric|openml|relbench|open-tabfm|frozen)" ; nvidia-smi --query-gpu=name,memory.total --format=csv,noheader; } > reports/J0_env.md
commit_push_sandbox "J0: env + weights + tests (job ${SLURM_JOB_ID})" reports
echo "=== J0 done $(date) ==="
