# TIMEOUT=1800
# rerun of 050 inside the project venv (050 ran in the runner conda env, which has no relbench)
# J18 rel-event/user-repeat: Kumo Relational's own prediction is at chance (0.49 on probe-train and test) while a linear
# probe on its graph layer reaches 0.79. Check whether that is the model or our score extraction: print the raw output
# columns and per-column AUROC on 256 test rows (CPU), with rel-f1/driver-dnf (own prediction 0.77, worked) as control.
source "$HOME/venvs/tabfm/bin/activate"; export HF_HUB_DOWNLOAD_TIMEOUT=60 OMP_NUM_THREADS=8; python -c "import relbench, sdm; print(\"env ok\")"
cd frozen-embeddings-retrieval
python - <<'PY'
import numpy as np, pandas as pd, torch, relbench, warnings
warnings.filterwarnings("ignore")
from sklearn.metrics import roc_auc_score
import fer.relbench_layers as RL
from sdm.models import KumoRelational
seen = {}
orig = RL._positive_scores
def spy(out, n):
    key = next(k for k in out.columns if str(k).endswith("numerical"))
    names = [str(c) for c in out.columns[key]]
    raw = torch.as_tensor(out.numerical).float().cpu().numpy()
    seen.setdefault("cols", (list(out.columns), names, raw.shape))
    seen.setdefault("raw", []).append(raw.reshape(-1, len(names))[-n:])
    return orig(out, n)
RL._positive_scores = spy
for ds, tn in [("rel-f1", "driver-dnf"), ("rel-event", "user-repeat")]:
    seen.clear()
    task = relbench.load_dataset(ds).load_task(tn)
    db = task.get_db(upto_test_timestamp=False)
    ent, tcol, tgt = task.entity_col, task.time_col, task.target_col
    tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
    te = task.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
    sub = RL.Subgrapher(db, task.entity_table, k_children=50)
    early, late, info = RL.time_split(tr, tcol, 4000)
    trs = early.sample(frac=1.0, random_state=0).drop_duplicates(ent)
    pos, neg = trs[trs[tgt] == 1], trs[trs[tgt] != 1]
    nc_pos = min(len(pos), max(256, int(512 * len(pos) / max(len(trs), 1))))
    ctx = pd.concat([pos.head(nc_pos), neg.head(512 - nc_pos)]).reset_index(drop=True)
    torch.manual_seed(0)
    model = KumoRelational(device="cpu"); model.eval()
    rows = te.iloc[np.random.default_rng(0).choice(len(te), min(256, len(te)), replace=False)].reset_index(drop=True)
    _, sc = RL.embed_rows(model, sub, ctx, ctx[tgt].to_numpy().astype(int), rows, ent, tcol, "cpu", batch=256)
    y = rows[tgt].astype(int).to_numpy()
    print(f"== {ds}/{tn} ctx={len(ctx)} ctx_pos={int(ctx[tgt].sum())} rows={len(rows)} rows_pos={int(y.sum())}")
    print("   columns:", seen["cols"][0], "numerical names:", seen["cols"][1], "raw shape:", seen["cols"][2])
    R = np.concatenate(seen["raw"])
    for j, nm in enumerate(seen["cols"][1]):
        c = R[:, j]
        print(f"   col {nm!r}: min {c.min():.4f} max {c.max():.4f} std {c.std():.4f} uniq {len(np.unique(c.round(5)))} auroc {roc_auc_score(y, c):.3f}")
    print(f"   extracted score auroc {roc_auc_score(y, sc):.3f}")
PY
