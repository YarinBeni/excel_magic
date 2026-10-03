# TIMEOUT=600
# J16: Kumo-S is nondeterministic (+-30% per split) but J2/J8/J16 agree; J15's p0 is systematically worse. Find what J15 did differently.
cd open-tabfm-auto
python - <<'PY'
import json, hashlib, glob
from pathlib import Path
from tabfm_auto.pipeline import IDENTITY_PIPELINE
from tabfm_auto.benchmarks.tabarena import list_datasets, load_openml_task, official_split
from tabfm_auto.harness.evaluator import evaluate_holdout
run = Path(sorted(glob.glob("../artifacts/runs/*J2_heuristic_anneal"))[-1])
ws = run / "search" / "anneal" / "workspace"
e1 = (ws / "candidates" / "eval_001.py").read_text()
print("eval_001 == IDENTITY:", e1 == IDENTITY_PIPELINE, hashlib.sha1(e1.encode()).hexdigest()[:8], hashlib.sha1(IDENTITY_PIPELINE.encode()).hexdigest()[:8])
if e1 != IDENTITY_PIPELINE:
    import difflib; print("".join(list(difflib.unified_diff(IDENTITY_PIPELINE.splitlines(1), e1.splitlines(1)))[:60]))
cfg = json.loads((run / "config.json").read_text()); print("cfg model:", cfg.get("model"), "max_rows:", cfg.get("max_rows"))
d = {x.name: x for x in list_datasets()}["anneal"]; task, oml = load_openml_task(d)
tr, te = official_split(oml, 0, 0)
Xtr, ytr, Xte, yte = task.X.iloc[tr].reset_index(drop=True), task.y.iloc[tr].reset_index(drop=True), task.X.iloc[te].reset_index(drop=True), task.y.iloc[te].reset_index(drop=True)
for label, p in [("workspace eval_001", ws / "candidates" / "eval_001.py"), ("fresh identity", Path("/tmp/j17_identity.py"))]:
    if label.startswith("fresh"): p.write_text(IDENTITY_PIPELINE)
    r = evaluate_holdout(p, Xtr, ytr, Xte, yte, task.task_type, cfg.get("model", "tabpfn"), seed=0, max_rows=cfg.get("max_rows", 10000))
    print(label, "->", r.get("score"), r.get("info"), r.get("error"))
# a pipeline listing of the workspace, in case the identity pipeline imports something local
print("workspace files:", sorted(x.name for x in ws.iterdir()))
PY
