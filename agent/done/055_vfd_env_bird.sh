# TIMEOUT=3600
# 054 showed `pip install gliclass gliner` upgraded transformers 4.57.6 -> 5.x inside the shared tabfm venv (breaks
# sentence-transformers 3.0.1, and GLiClass has a known transformers-5 bug). (1) restore the shared venv, (2) build a
# separate venv for the vfd study, (3) download the 11 BIRD dev databases (official mirror first).
set +e
echo "== 1. restore tabfm venv"
source "$HOME/venvs/tabfm/bin/activate"
pip install -q "transformers==4.57.6" 2>&1 | grep -v "^WARNING" | tail -2
python -c "import transformers, torch, sentence_transformers; print('tabfm: transformers', transformers.__version__, 'torch', torch.__version__, 'st', sentence_transformers.__version__)"
deactivate
echo "== 2. vfd venv"
PYBIN="$HOME/miniconda3/envs/thesis/bin/python"
[ -x "$HOME/venvs/vfd/bin/python" ] || "$PYBIN" -m venv "$HOME/venvs/vfd"
source "$HOME/venvs/vfd/bin/activate"
pip install -q --upgrade pip 2>&1 | tail -1
pip install -q "torch==2.9.1" --index-url https://download.pytorch.org/whl/cu128 2>&1 | tail -2
pip install -q "transformers>=4.51,<5" gliclass gliner "sentence-transformers>=3,<5" sqlglot openai scikit-learn pandas tabulate requests huggingface_hub pytest 2>&1 | grep -iv "^warning" | tail -4
python -c "import torch, transformers, gliclass, gliner, sentence_transformers, sqlglot, openai; print('vfd: torch', torch.__version__, torch.version.cuda, 'transformers', transformers.__version__, 'gliner', gliner.__version__, 'st', sentence_transformers.__version__)"
echo "== 3. BIRD dev databases"
mkdir -p artifacts/bird
python - <<'PY'
import sqlite3
from pathlib import Path
from huggingface_hub import list_repo_files, snapshot_download
need = ["california_schools", "card_games", "codebase_community", "debit_card_specializing", "european_football_2",
        "financial", "formula_1", "student_club", "superhero", "thrombosis_prediction", "toxicology"]
root = Path("artifacts/bird")
for rid in ["birdsql/bird_sql_dev_20251106", "premai-io/birdbench", "Open-Dataflow/dataflow-Text2SQL-database-example",
            "xu3kev/BIRD-SQL-data"]:
    missing = [d for d in need if not list(root.rglob(f"{d}.sqlite"))]
    if not missing:
        break
    try:
        fs = list_repo_files(rid, repo_type="dataset")
    except Exception as e:
        print("LIST", rid, "ERR", str(e)[:120]); continue
    hits = [f for f in fs if any(f"/{d}/" in f"/{f}" for d in missing)]
    zips = [f for f in fs if f.endswith(".zip")]
    print("REPO", rid, len(fs), "files;", len(hits), "matching paths; zips:", zips[:6])
    if hits:
        pats = [f"*{d}/*" for d in missing] + [f"*{d}/database_description/*" for d in missing]
        snapshot_download(rid, repo_type="dataset", allow_patterns=pats, local_dir=str(root / rid.replace("/", "__")))
    elif zips:
        z = [f for f in zips if "dev" in f.lower()][:1]
        if z:
            import zipfile
            p = snapshot_download(rid, repo_type="dataset", allow_patterns=z, local_dir=str(root / rid.replace("/", "__")))
            for f in Path(p).rglob("*.zip"):
                zipfile.ZipFile(f).extractall(f.parent)
                for inner in f.parent.rglob("dev_databases.zip"):
                    zipfile.ZipFile(inner).extractall(inner.parent)
for d in need:
    h = sorted(root.rglob(f"{d}.sqlite"))
    if not h:
        print("MISSING", d); continue
    con = sqlite3.connect(h[0]); n = con.execute("select count(*) from sqlite_master where type='table'").fetchone()[0]
    desc = len(list((h[0].parent / "database_description").glob("*.csv")))
    print("DB", d, h[0], f"{h[0].stat().st_size/1e6:.0f} MB", n, "tables", desc, "description csvs")
PY
du -sh artifacts/bird
