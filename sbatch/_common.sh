# _common.sh — prologue shared by every tabfm job. Source AFTER cd into the repo root.
#   source sbatch/_common.sh && job_prologue   (sync branch, env, log push, exported paths)
REPO="$HOME/projects/excel_magic"
VENV="$HOME/venvs/tabfm"

job_prologue() {
    source ~/miniconda3/etc/profile.d/conda.sh
    conda activate thesis
    cd "$REPO" || { echo "FAILED: $REPO missing"; exit 1; }
    source sbatch/_sandbox.sh
    sync_sandbox || { echo "FAILED sync_sandbox: refusing to run stale code"; exit 1; }
    enable_log_push
    mkdir -p sbatch/logs artifacts/runs artifacts/weights artifacts/openml reports/runs
    export TABFM_WEIGHTS_DIR="$REPO/artifacts/weights"
    export TABFM_RUNS_ROOT="$REPO/artifacts/runs"
    export OPENML_CACHE_DIR="$REPO/artifacts/openml"
    export HF_HUB_DOWNLOAD_TIMEOUT=60
    export TOKENIZERS_PARALLELISM=false
    export FER_DEVICE=cuda
    export OMP_NUM_THREADS="${SLURM_CPUS_PER_TASK:-4}"
    if [ -f "$VENV/bin/activate" ]; then source "$VENV/bin/activate"; fi
    echo "python=$(which python) $(python --version 2>&1) gpu=$(nvidia-smi --query-gpu=name --format=csv,noheader 2>/dev/null | head -1)"
}

# Copy the small, reviewable artifacts of every run under artifacts/runs into tracked reports/runs and push.
collect_and_push() {
    local msg="$1"
    python3 - <<'PY'
import os, shutil, pathlib
src = pathlib.Path("artifacts/runs"); dst = pathlib.Path("reports/runs"); dst.mkdir(parents=True, exist_ok=True)
keep = {"metrics.json", "config.json", "results.md", "results.csv", "best_pipeline.py", "build_table.sql", "agent_result.md",
        "comparison_to_paper.md", "comparison_to_paper.csv", "heuristic_trace.json", "rescored.json", "NOTES.md",
        "agent_stream.log", "agent_stream.jsonl", "agent_result.md", "relbench_rows.json", "churn_rows.json",
        "selection_rules.md", "selection_rules.csv", "layers_rows.json",
        "candidates.jsonl", "signals.csv", "schema_link.csv", "judge.csv", "encoders.csv", "rows.jsonl", "scores.jsonl", "harness_rows.jsonl", "capability_rows.jsonl"}
n = 0
for run in sorted(src.glob("*")):
    if not run.is_dir(): continue
    for f in run.rglob("*"):
        if f.name.startswith("agent_stream") and not (f.parents[0] / "metrics.json").exists() \
                and not (run / "metrics.json").exists():
            continue  # a growing log of a run still in progress: snapshotting it twice makes add/add conflicts
        if f.is_file() and (f.name in keep or f.name.startswith(("labels_", "insights_"))) and f.stat().st_size < 5_000_000:
            out = dst / run.name / f.relative_to(run)
            out.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, out); n += 1
print(f"[collect] copied {n} files into reports/runs")
PY
    commit_push_sandbox "$msg" reports || echo "[collect] push failed; commit kept locally"
}
