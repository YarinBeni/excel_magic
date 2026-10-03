#!/bin/bash
#SBATCH --job-name=J16_determinism
#SBATCH --partition=voltagepark
#SBATCH --gres=gpu:1
#SBATCH --cpus-per-task=8
#SBATCH --mem=64G
#SBATCH --time=01:00:00
#SBATCH --output=sbatch/logs/J16_determinism_%j.out
# J16 — why does the identity pipeline score differently in J15 (23:35) than in J2 (15:50) / J8 (19:20) on the same splits?
# Score P0 on anneal / airfoil r0f0..r0f2 three times in one process and in a second process, with the CURRENT library and
# with the library as of the J8-era commit (3b9c16a), printing sdm/torch versions. Writes reports/J16_determinism.md.
set -uo pipefail
echo "=== J16 determinism === job ${SLURM_JOB_ID} $(date)"
cd ~/projects/excel_magic && source sbatch/_common.sh && job_prologue
OUT=reports/J16_determinism.md; echo "# J16 determinism ($(date), job ${SLURM_JOB_ID})" > $OUT
cat > /tmp/j16_score_$$.py <<'PY'
import sys, json, numpy as np
from tabfm_auto.benchmarks.tabarena import list_datasets, load_openml_task, official_split
from tabfm_auto.harness.evaluator import evaluate_holdout
from tabfm_auto.pipeline import IDENTITY_PIPELINE
import importlib.metadata as im, torch
from pathlib import Path
tag = sys.argv[1]
print(f"[{tag}] torch {torch.__version__} sdm {im.version('structured-data-models') if True else '?'} cuda {torch.cuda.is_available()}", flush=True)
p0 = Path("/tmp/j16_identity.py"); p0.write_text(IDENTITY_PIPELINE)
ds = {d.name: d for d in list_datasets()}
for name in ["anneal", "airfoil_self_noise"]:
    task, oml = load_openml_task(ds[name])
    for f in range(3):
        tr, te = official_split(oml, 0, f)
        Xtr, ytr = task.X.iloc[tr].reset_index(drop=True), task.y.iloc[tr].reset_index(drop=True)
        Xte, yte = task.X.iloc[te].reset_index(drop=True), task.y.iloc[te].reset_index(drop=True)
        scores = [evaluate_holdout(p0, Xtr, ytr, Xte, yte, task.task_type, "kumo-tabular-s:n_estimators=8,device=cuda", seed=f)["score"] for _ in range(3)]
        print(f"[{tag}] {name} r0f{f}: " + " ".join(f"{s:.5f}" for s in scores), flush=True)
PY
cd open-tabfm-auto
echo "## current library ($(git -C .. rev-parse --short HEAD))" >> ../$OUT
python /tmp/j16_score_$$.py current 2>&1 | grep "^\[current\]" | tee -a ../$OUT
python /tmp/j16_score_$$.py current-proc2 2>&1 | grep "^\[current-proc2\]" | tee -a ../$OUT
# the J8-era library in a scratch worktree (same venv, PYTHONPATH override)
rm -rf /tmp/j16_old_$$ && git -C .. worktree add -q /tmp/j16_old_$$ 3b9c16a 2>/dev/null || git -C .. worktree add -q -f /tmp/j16_old_$$ 3b9c16a
echo "## J8-era library (3b9c16a)" >> ../$OUT
( cd /tmp/j16_old_$$/open-tabfm-auto && PYTHONPATH=/tmp/j16_old_$$/open-tabfm-auto python /tmp/j16_score_$$.py old 2>&1 | grep "^\[old\]" ) | tee -a ../$OUT
git -C .. worktree remove -f /tmp/j16_old_$$ 2>/dev/null || true
cd ..
# J2 / J8 reference numbers for the same splits
python - <<'PY' | tee -a $OUT
import pandas as pd, glob
print("## reference: P0 from the original runs (r0f0..r0f2)")
for tag,pat in [("J2 heuristic","*J2_heuristic_%s"),("J8 pi","*J8_pi_qwen3coder_%s"),("J15 rescore(p0)","*J2_heuristic_%s")]:
    for name in ["anneal","airfoil_self_noise"]:
        fs=glob.glob(f"reports/runs/{pat % name}/"+("selection_rules.csv" if "J15" in tag else "results.csv"))
        if not fs: continue
        d=pd.read_csv(fs[0]); d=d[(d.repeat==0)]
        d=d[d.rule=="p0"] if "J15" in tag else d[d.pipeline=="P0"]
        print(f"[{tag}] {name}: " + " ".join(f"{v:.5f}" for v in d.sort_values('fold').error))
PY
commit_push_sandbox "J16: determinism check (job ${SLURM_JOB_ID})" reports
echo "=== J16 done $(date) ==="
