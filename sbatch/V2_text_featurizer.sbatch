#!/bin/bash
#SBATCH --job-name=V2_text_featurizer
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=12:00:00
#SBATCH --output=sbatch/logs/V2_text_featurizer_%j.out
# V2 (phase P1, claim C5): GLiClass labels of free-text columns vs tf-idf / embeddings as features for frozen Kumo
# Tabular-L, TabPFN-2.5 and LightGBM, on the 18-dataset AutoGluon multimodal text+tabular benchmark (official splits).
set -uo pipefail
echo "=== V2 text featurizer === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
RUN="$TABFM_RUNS_ROOT/$(date -u +%Y%m%dT%H%M%SZ)_V2_text_featurizer"; mkdir -p "$RUN"
export MMBENCH_DATA="$REPO/artifacts/mmbench" AUTO_MM_BENCH_HOME="$REPO/artifacts/mmbench_raw"
deactivate 2>/dev/null || true; source "$HOME/venvs/vfd/bin/activate"
python -c "import auto_mm_bench" 2>/dev/null || {
  pip install -q boto3 tqdm pyarrow 2>&1 | tail -1
  pip install -q --no-deps "git+https://github.com/sxjscience/automl_multimodal_benchmark.git#subdirectory=multimodal_text_benchmark" 2>&1 | tail -1; }
P="python vfd-agent/experiments/p1_text_featurizer.py --run $RUN ${DATASETS:+--datasets $DATASETS}"
echo "--- prepare $(date)"; $P prepare || { echo "FAILED prepare"; collect_and_push "V2: prepare failed (job ${SLURM_JOB_ID})"; exit 1; }
echo "--- labels $(date)"
source sbatch/_vllm.sh
LLM=Qwen/Qwen3-Coder-30B-A3B-Instruct PORT=8210 VLLM_GPU_FRAC=0.85 VLLM_MAX_LEN=32768 start_vllm "qwen3_coder hermes" \
  && { OPENAI_BASE_URL=http://localhost:8210/v1 OPENAI_API_KEY=EMPTY $P labels || echo "FAILED labels"; } || echo "FAILED vllm"
kill $VLLM_PID 2>/dev/null; sleep 20
echo "--- features $(date)"; $P features || echo "FAILED features"
echo "--- evaluate $(date)"
deactivate; source "$HOME/venvs/tabfm/bin/activate"
$P evaluate || echo "FAILED evaluate"
$P report || echo "FAILED report"
cp "$MMBENCH_DATA"/*/labels.json "$RUN"/ 2>/dev/null; for d in "$MMBENCH_DATA"/*/; do n=$(basename $d); [ -f "$d/labels.json" ] && cp "$d/labels.json" "$RUN/labels_$n.json"; done
collect_and_push "V2: text featurizer (GLiClass labels vs embeddings) on 18 text+tabular datasets (job ${SLURM_JOB_ID})"
echo "=== V2 done $(date) ==="
