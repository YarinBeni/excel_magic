# TIMEOUT=1500
# P0 setup probe for the three-model harness study: install GLiNER / GLiClass into the project venv, find a reachable
# source for the BIRD mini-dev SQLite databases, check that the verifier models are downloadable.
source "$HOME/venvs/tabfm/bin/activate"
python -c "import transformers, torch; print('transformers', transformers.__version__, 'torch', torch.__version__)"
pip install -q gliclass gliner sqlglot 2>&1 | tail -3
python -c "import gliclass, gliner, sqlglot; print('gliclass ok', 'gliner', gliner.__version__, 'sqlglot', sqlglot.__version__)" 2>&1 | tail -3
for u in https://bird-bench.oss-cn-beijing.aliyuncs.com/minidev.zip https://bird-bench.oss-cn-beijing.aliyuncs.com/dev.zip; do
  echo "== $u"; curl -sI --max-time 20 "$u" | head -3; done
python - <<'PY'
from huggingface_hub import list_repo_files, model_info
for rid in ["birdsql/bird_mini_dev", "heegyu/bird-sql-mini-dev", "567-labs/bird-dev-tables", "epiphanian/bird_mini_dev"]:
    try:
        fs = list_repo_files(rid, repo_type="dataset")
        big = [f for f in fs if f.endswith((".sqlite", ".zip", ".tar.gz", ".db"))]
        print("DATASET", rid, len(fs), "files; db-like:", big[:15])
    except Exception as e:
        print("DATASET", rid, "ERR", type(e).__name__, str(e)[:120])
for m in ["knowledgator/gliclass-large-v3.0", "knowledgator/gliclass-base-v3.0", "knowledgator/gliner-bi-base-v2.0",
          "urchade/gliner_large-v2.1", "MoritzLaurer/deberta-v3-large-zeroshot-v2.0", "BAAI/bge-reranker-v2-m3",
          "lytang/MiniCheck-Flan-T5-Large", "BAAI/bge-small-en-v1.5"]:
    try:
        mi = model_info(m); print("MODEL", m, "ok", (mi.library_name or ""), mi.sha[:8])
    except Exception as e:
        print("MODEL", m, "ERR", type(e).__name__, str(e)[:100])
PY
ls ~/miniconda3/envs/vllm/.ok 2>/dev/null && echo "vllm env ok"
