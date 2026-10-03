# TIMEOUT=1800
# BIRD column descriptions live inside the official dev_databases.zip (mirrored at Sudnya/bird-sql): extract only
# */database_description/*.csv into the premai copy of the 11 dev databases.
set +e
source "$HOME/venvs/vfd/bin/activate"
python - <<'PY'
import zipfile, shutil
from pathlib import Path
from huggingface_hub import hf_hub_download
dst = Path("artifacts/bird/premai-io__birdbench/validation/dev_databases")
z = hf_hub_download("Sudnya/bird-sql", "databases/dev_databases.zip", repo_type="dataset", local_dir="artifacts/bird/sudnya")
n = 0
with zipfile.ZipFile(z) as zf:
    for m in zf.namelist():
        parts = Path(m).parts
        if "database_description" in parts and m.lower().endswith(".csv") and not parts[-1].startswith("._"):
            db = parts[parts.index("database_description") - 1]
            if (dst / db).exists():
                out = dst / db / "database_description" / parts[-1]
                out.parent.mkdir(exist_ok=True)
                out.write_bytes(zf.read(m)); n += 1
print("extracted", n, {p.name: len(list((p / "database_description").glob("*.csv"))) for p in dst.iterdir() if p.is_dir()})
Path(z).unlink()
PY
