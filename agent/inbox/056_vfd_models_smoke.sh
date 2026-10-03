# TIMEOUT=2400
# (1) BIRD column descriptions (database_description/*.csv) from a second mirror into the premai copy;
# (2) functional CPU smoke of every small model in the vfd venv (GLiClass declares transformers>=5, venv has 4.57.6).
set +e
source "$HOME/venvs/vfd/bin/activate"
python - <<'PY'
import shutil
from pathlib import Path
from huggingface_hub import snapshot_download
dst = Path("artifacts/bird/premai-io__birdbench/validation/dev_databases")
src = Path(snapshot_download("Open-Dataflow/dataflow-Text2SQL-database-example", repo_type="dataset",
                             allow_patterns=["*dev_databases/*/database_description/*"], local_dir="artifacts/bird/dataflow_desc"))
n = 0
for d in src.rglob("database_description"):
    db = d.parent.name
    if (dst / db).exists():
        (dst / db / "database_description").mkdir(exist_ok=True)
        for f in d.glob("*.csv"):
            shutil.copy2(f, dst / db / "database_description" / f.name); n += 1
print("copied", n, "description csvs;", {p.name: len(list((p / 'database_description').glob('*.csv'))) for p in dst.iterdir() if p.is_dir()})
PY
echo "== smoke"
cat > /tmp/vfd_smoke_$$.py <<'PY'
import time, traceback
texts = ["Question: How many customers pay in EUR?\nSQL: SELECT count(*) FROM customers WHERE Currency = 'EUR'\nResult: 1 rows: 23"]
def t(name, f):
    t0 = time.time()
    try:
        print("OK", name, f(), f"{time.time()-t0:.1f}s", flush=True)
    except Exception as e:
        print("FAIL", name, type(e).__name__, str(e)[:300], flush=True); traceback.print_exc(limit=2)
def gliclass_test():
    from gliclass import GLiClassModel, ZeroShotClassificationPipeline
    from transformers import AutoTokenizer
    m = GLiClassModel.from_pretrained("knowledgator/gliclass-base-v3.0")
    tok = AutoTokenizer.from_pretrained("knowledgator/gliclass-base-v3.0", add_prefix_space=True)
    p = ZeroShotClassificationPipeline(m, tok, classification_type="multi-label", device="cpu")
    r = p(texts * 2, ["The SQL query correctly answers the question.", "wrong aggregation"], threshold=0.0)
    return type(r).__name__, len(r), r[0]
def gliner_test():
    from gliner import GLiNER
    m = GLiNER.from_pretrained("knowledgator/gliner-bi-base-v2.0")
    return m.predict_entities("customers who pay in EUR in 2012", ["customers currency", "transaction date"], threshold=0.05)
def nli_test():
    from transformers import pipeline
    p = pipeline("zero-shot-classification", model="MoritzLaurer/deberta-v3-large-zeroshot-v2.0", device=-1)
    return p(texts, candidate_labels=["The SQL query correctly answers the question."], hypothesis_template="{}", multi_label=True)
def rerank_test():
    from sentence_transformers import CrossEncoder, SentenceTransformer
    c = CrossEncoder("BAAI/bge-reranker-v2-m3", device="cpu")
    e = SentenceTransformer("BAAI/bge-small-en-v1.5", device="cpu")
    return c.predict([("pay in EUR", "customers.Currency: currency"), ("pay in EUR", "yearmonth.Date")]).tolist(), e.encode(["x"]).shape
for n, f in [("gliclass", gliclass_test), ("gliner", gliner_test), ("nli", nli_test), ("rerank+embed", rerank_test)]:
    t(n, f)
PY
python /tmp/vfd_smoke_$$.py 2>&1 | grep -v "Warning\|warn(" | tail -30
pip list 2>/dev/null | grep -iE "^(gliclass|gliner|transformers|sentence-transformers|torch) "
