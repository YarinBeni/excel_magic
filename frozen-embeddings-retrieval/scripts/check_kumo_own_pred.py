"""Is Kumo Relational's own in-context prediction really at chance on a task, or is our score extraction wrong?

For each task: rebuild the J18 time-split context (real labels), score 512 test rows, and print the raw output columns
(per-class probabilities), their spread, and the AUROC of each column against the labels, next to the score that
``fer.relbench_layers.embed_rows`` extracts. Rows are sorted by time first, so the batch order equals the row order.

  python scripts/check_kumo_own_pred.py --tasks rel-f1/driver-dnf,rel-event/user-repeat --device cuda
"""
from __future__ import annotations

import argparse
import warnings

import numpy as np
import pandas as pd


def main() -> None:
    import relbench
    import torch
    from sdm.models import KumoRelational
    from sklearn.metrics import roc_auc_score

    import fer.relbench_layers as RL

    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", default="rel-f1/driver-dnf,rel-event/user-repeat")
    ap.add_argument("--device", default="cuda")
    ap.add_argument("--n", type=int, default=512)
    a = ap.parse_args()
    warnings.filterwarnings("ignore")
    raw_out: list[np.ndarray] = []
    meta: dict = {}
    orig = RL._positive_scores

    def spy(out, n):
        key = next(k for k in out.columns if str(k).endswith("numerical"))
        names = [str(c) for c in out.columns[key]]
        meta.setdefault("names", names)
        raw_out.append(torch.as_tensor(out.numerical).float().cpu().numpy().reshape(-1, len(names))[-n:])
        return orig(out, n)

    RL._positive_scores = spy
    for spec in a.tasks.split(","):
        ds, tn = spec.split("/")
        raw_out.clear(); meta.clear()
        task = relbench.load_dataset(ds).load_task(tn)
        db = task.get_db(upto_test_timestamp=False)
        ent, tcol, tgt = task.entity_col, task.time_col, task.target_col
        tr = task.get_table("train", mask_input_cols=False).df.reset_index(drop=True)
        te = task.get_table("test", mask_input_cols=False).df.reset_index(drop=True)
        sub = RL.Subgrapher(db, task.entity_table, k_children=50)
        early, _, info = RL.time_split(tr, tcol, 4000)
        trs = early.sample(frac=1.0, random_state=0).drop_duplicates(ent)
        pos, neg = trs[trs[tgt] == 1], trs[trs[tgt] != 1]
        nc_pos = min(len(pos), max(256, int(512 * len(pos) / max(len(trs), 1))))
        ctx = pd.concat([pos.head(nc_pos), neg.head(512 - nc_pos)]).reset_index(drop=True)
        torch.manual_seed(0)
        model = KumoRelational(device=a.device)
        model.eval()
        rows = te.iloc[np.random.default_rng(0).choice(len(te), min(a.n, len(te)), replace=False)]
        rows = rows.sort_values(tcol, kind="stable").reset_index(drop=True)
        _, sc = RL.embed_rows(model, sub, ctx, ctx[tgt].to_numpy().astype(int), rows, ent, tcol, a.device, batch=256)
        y = rows[tgt].astype(int).to_numpy()
        R = np.concatenate(raw_out)
        print(f"== {spec}  target values {dict(tr[tgt].value_counts())}  ctx={len(ctx)} ctx_pos={int(ctx[tgt].sum())}  "
              f"rows={len(rows)} rows_pos={int(y.sum())}  {info}")
        print(f"   output columns {meta['names']}  raw {R.shape}")
        for j, nm in enumerate(meta["names"]):
            c = R[:, j]
            print(f"   col {nm!r}: min {c.min():.4f} max {c.max():.4f} std {c.std():.5f} distinct {len(np.unique(c.round(5)))} "
                  f"auroc {roc_auc_score(y, c):.3f}")
        print(f"   extracted score: auroc {roc_auc_score(y, sc):.3f}  std {sc.std():.5f}", flush=True)


if __name__ == "__main__":
    main()
