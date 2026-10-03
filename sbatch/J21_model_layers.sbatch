#!/bin/bash
#SBATCH --job-name=J21_model_layers
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=16
#SBATCH --mem=160G
#SBATCH --time=10:00:00
#SBATCH --output=sbatch/logs/J21_model_layers_%j.out
# J21 — every open tabular foundation model (TabPFN v2, TabPFN-2.5, TabICLv2, Kumo Tabular S/M/L), every layer, on 5 real
# RelBench entity tasks: linear/kNN probes (layer chosen on val, official test evaluator), label-free geometry per layer
# (effective rank, TwoNN intrinsic dimension, anisotropy, k-means silhouette), CKA between layers and between models.
set -uo pipefail
echo "=== J21 model layers === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
cd open-tabfm-auto && tabfm-models download tabpfn-2.5 tabiclv2-sdm kumo-tabular-m kumo-tabular-l 2>&1 | tail -4; tabfm-models status 2>&1 | grep -E "tabpfn|tabicl|kumo" | head -10; cd ..
cd frozen-embeddings-retrieval
mkdir -p weights && ln -sfn "$TABFM_WEIGHTS_DIR/tabpfn" weights/tabpfn
# smoke: one small task, two models, one target; every tapped layer must vary across rows
python experiments/exp06_model_layers.py --tasks rel-f1/driver-dnf --models tabpfn,kumo-s --targets label --n-ctx 500 --n-train 500 --name J21_smoke \
    || { echo "FAILED smoke"; cd ..; collect_and_push "J21: smoke failed (job ${SLURM_JOB_ID})"; exit 1; }
python - <<'PY' || { echo "FAILED smoke check"; cd ..; collect_and_push "J21: smoke check failed (job ${SLURM_JOB_ID})"; exit 1; }
import glob, json, os
f = sorted(glob.glob(os.path.join(os.environ["TABFM_RUNS_ROOT"], "*J21_smoke*", "layers_rows.json")))[-1]
r = json.load(open(f))
for mname, per in r["models"].items():
    for t, m in per.items():
        assert "error" not in m, (mname, t, m.get("error"))
        bad = [L for L, v in m["layers"].items() if v["geometry"]["effective_rank"] < 1.5]
        print(mname, t, "layers", len(m["layers"]), "degenerate", bad, "best", m["best_by_val_linear"])
        assert not bad, f"degenerate layers {bad}"
PY
TASKS="${TASKS:-rel-f1/driver-dnf,rel-trial/study-outcome,rel-f1/driver-top3,rel-event/user-repeat,rel-event/user-ignore}"
IFS=',' read -ra TL <<< "$TASKS"
for T in "${TL[@]}"; do
    echo "--- $T $(date)"
    python experiments/exp06_model_layers.py --tasks "$T" --name J21_model_layers || echo "[J21] $T FAILED"
    ( cd .. && collect_and_push "J21: all-model layer sweep $T (job ${SLURM_JOB_ID})" )
done
echo "=== J21 done $(date) ==="
