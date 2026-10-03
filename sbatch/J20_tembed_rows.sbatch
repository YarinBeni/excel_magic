#!/bin/bash
#SBATCH --job-name=J20_tembed_rows
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=96G
#SBATCH --time=07:00:00
#SBATCH --output=sbatch/logs/J20_tembed_rows_%j.out
# J20 — TEmBed (IBM, 2026) row-level embedding benchmark: row similarity search (entity-resolution datasets, MAP) and
# row triplet evaluation (Wikidata), TabPFN v2 with the context target / layer as the only change (zeros = TEmBed's
# default), plus the MiniLM text-embedding reference. Results compared with the TEmBed paper's published tabpfn row.
set -uo pipefail
echo "=== J20 tembed rows === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
R=$HOME/tembed; OUT=reports/J20_tembed; mkdir -p $OUT
# the repo holds a template folder literally named "<approach_name>", which this filesystem refuses: sparse-checkout
# everything except it (and drop a half-checked-out clone left by an earlier attempt)
[ -f $R/run_experiments.py ] || rm -rf "${R:?}"
if [ ! -d $R/.git ]; then
    git clone -q --no-checkout https://github.com/IBM/table-representation-evals.git $R || { echo "FAILED clone"; exit 1; }
    git -C $R sparse-checkout init --no-cone
    git -C $R sparse-checkout set --no-cone '/*' '!/approaches/benchmark_approaches_src/<approach_name>/'
    git -C $R checkout -q || { echo "FAILED checkout"; exit 1; }
fi
ls $R | head -20
cp -r tembed_patch/approaches/benchmark_approaches_src/tabpfn_ctx $R/approaches/benchmark_approaches_src/
cp tembed_patch/configs/approaches/tabpfn_ctx.yaml $R/configs/approaches/
cp tembed_patch/configs/runs/tabpfn_ctx_rows.yaml tembed_patch/configs/runs/tabpfn_ctx_rowsim.yaml $R/configs/runs/
TABPFN_VER=$(python -c "import importlib.metadata as m; print(m.version('tabpfn'))")
source ~/miniconda3/etc/profile.d/conda.sh
( flock -w 3600 8 || exit 1
  if ! conda env list | grep -q '^benchmark_env '; then
      conda create -y -q -n benchmark_env python=3.12 pip || exit 1
      conda activate benchmark_env
      cd $R && pip install -q -r reqs_benchmark.txt && pip install -q torch==2.9.1 --index-url https://download.pytorch.org/whl/cu128 \
          && pip install -q "tabpfn==$TABPFN_VER" && pip install -q -e . -e approaches/ || exit 1
  fi
) 8>$HOME/.tembed_env.lock || { echo "FAILED env"; cd ~/projects/excel_magic; collect_and_push "J20: env failed (job ${SLURM_JOB_ID})"; exit 1; }
conda activate benchmark_env
cd $R
if [ ! -d cache/row_similarity_data_full ] || [ -z "$(ls -A cache/row_similarity_data_full 2>/dev/null)" ]; then
    mkdir -p cache/raw_datasets   # the creation script downloads into it but asserts it exists first
    bash ./benchmark_src/dataset_creation/em_datasets/create_benchmark.sh 2>&1 | tail -20 || echo "[J20] em dataset creation reported errors"
fi
ls cache/row_similarity_data_full 2>/dev/null | head -20 | tee ~/projects/excel_magic/$OUT/datasets.txt
export TABPFN_CKPT="$TABFM_WEIGHTS_DIR/tabpfn/tabpfn-v2-classifier.ckpt" PYTHONPATH="$R:${PYTHONPATH:-}"
python run_experiments.py "${RUN_CONFIG:-tabpfn_ctx_rows}" 2>&1 | tail -60 | tee ~/projects/excel_magic/$OUT/run_tail.txt
# collect every results.json into one table
python - <<'PY' | tee ~/projects/excel_magic/$OUT/summary.md
import json, glob, os
rows = []
for f in glob.glob("results/tabpfn_ctx_rows/**/results.json", recursive=True):
    parts = f.split(os.sep)
    try:
        r = json.load(open(f))
    except Exception:
        continue
    flat = {k: v for k, v in r.items() if isinstance(v, (int, float))}
    rows.append(("/".join(parts[2:-1]), flat))
print("| run (approach/params/task/dataset) | metrics |"); print("|---|---|")
for k, v in sorted(rows):
    print(f"| {k} | " + ", ".join(f"{a}={b:.4f}" for a, b in sorted(v.items())[:6]) + " |")
PY
mkdir -p ~/projects/excel_magic/$OUT/results && cp -r results/tabpfn_ctx_rows ~/projects/excel_magic/$OUT/results/ 2>/dev/null || true
find ~/projects/excel_magic/$OUT/results -size +2M -delete 2>/dev/null || true
cd ~/projects/excel_magic && commit_push_sandbox "J20: TEmBed row tasks, TabPFN context target / layer (job ${SLURM_JOB_ID})" reports
echo "=== J20 done $(date) ==="
