# TIMEOUT=900
# (1) V1 smoke on 30 questions (GPU job); (2) find BIRD database_description files on HF mirrors;
# (3) gliclass 0.1.17 (works with transformers 4.57.6) into the shared tabfm venv for the analyst agent.
set +e
J=$(sbatch --parsable --export=ALL,LIMIT=30 --time=03:00:00 sbatch/V1_verifier_study.sbatch); echo "$J" > agent/state/V1smoke.id; echo "V1 smoke: $J"
source "$HOME/venvs/vfd/bin/activate"
python - <<'PY'
from huggingface_hub import list_repo_files
for rid in ["Open-Dataflow/dataflow-Text2SQL-database-example", "xu3kev/BIRD-SQL-data", "premai-io/birdbench",
            "birdsql/bird_sql_dev_20251106", "Sudnya/bird-sql"]:
    try:
        fs = list_repo_files(rid, repo_type="dataset")
        d = [f for f in fs if "description" in f.lower()]
        print(rid, len(fs), "files;", len(d), "description paths;", d[:4], "| sample:", fs[:5])
    except Exception as e:
        print(rid, "ERR", str(e)[:100])
PY
deactivate
source "$HOME/venvs/tabfm/bin/activate"
pip install -q --no-deps "gliclass==0.1.17" 2>&1 | tail -1
python -c "import gliclass, transformers, tabfm_auto; print('tabfm venv: gliclass', gliclass.__version__ if hasattr(gliclass,'__version__') else 'ok', 'transformers', transformers.__version__)"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
