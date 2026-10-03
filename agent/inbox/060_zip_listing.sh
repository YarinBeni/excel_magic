# TIMEOUT=900
source "$HOME/venvs/vfd/bin/activate"
python - <<'PY'
import zipfile
from huggingface_hub import hf_hub_download
for rid, f in [("Sudnya/bird-sql", "databases/dev_databases.zip"), ("Open-Dataflow/dataflow-Text2SQL-database-example", "databases.zip")]:
    try:
        z = hf_hub_download(rid, f, repo_type="dataset", local_dir="artifacts/bird/ziplist")
        names = zipfile.ZipFile(z).namelist()
        d = [n for n in names if "desc" in n.lower()]
        print(rid, len(names), "members;", len(d), "desc members; e.g.", d[:5], "| first:", names[:8])
    except Exception as e:
        print(rid, "ERR", str(e)[:200])
PY
rm -rf artifacts/bird/ziplist
