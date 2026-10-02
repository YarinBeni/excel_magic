"""Export RelBench entity tasks as multi-table (N event tables) and run ICL eval
with the N-table database-native RFM. Leak-free: strict < anchor, unique-timestamp split.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
import pandas as pd
import torch
from sklearn.metrics import mean_absolute_error, roc_auc_score

from .multitable import MultiTableRFM
from .multitable_io import MultiTableDataset
from .relational_io import TASK_TYPES, temporal_train_end
from .benchmark_adapter import (
    relbench_task_type_to_adapter, prepare_table, prepare_task_table, find_fk_to_table,
    choose_children_by_coverage,
)
from .train import json_safe


def export_multitable(dataset, task_name, output_dir, max_tables=6, max_rows=1500000):
    from relbench.datasets import get_dataset
    from relbench.tasks import get_task
    ds = get_dataset(dataset, download=True)
    task = get_task(dataset, task_name, download=True)
    entity_table = getattr(task, "entity_table"); entity_col = getattr(task, "entity_col")
    time_col = getattr(task, "time_col"); target_col = getattr(task, "target_col")
    adapter_tt = relbench_task_type_to_adapter(getattr(task, "task_type", None))
    db = ds.get_db(upto_test_timestamp=True)
    task_df = task.get_table("train", mask_input_cols=False).df
    table_dict = db.table_dict
    root_src = table_dict[entity_table]
    root_df, root_pk, root_time, root_feats = prepare_table(root_src.df, root_src.pkey_col, root_src.time_col, root_src.fkey_col_to_pkey_table, max_rows)
    task_out = prepare_task_table(task_df, entity_col, time_col, target_col, adapter_tt)
    root_ids = set(root_df[root_pk].dropna().tolist())
    task_out = task_out[task_out[entity_col].isin(root_ids)].sort_values(time_col).reset_index(drop=True)

    task_entities = set(task_df[entity_col].dropna().tolist())
    child_names = choose_children_by_coverage(table_dict, entity_table, task_entities, top_k=max_tables)
    output_dir = Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    root_df.to_parquet(output_dir / "root.parquet", index=False)
    task_out.to_parquet(output_dir / "task.parquet", index=False)
    event_specs = []
    for i, cname in enumerate(child_names):
        csrc = table_dict[cname]
        cfk = find_fk_to_table(csrc.fkey_col_to_pkey_table, entity_table)
        cdf, cpk, ctime, cfeats = prepare_table(csrc.df, csrc.pkey_col, csrc.time_col, csrc.fkey_col_to_pkey_table, max_rows)
        cdf = cdf[cdf[cfk].isin(root_ids)].reset_index(drop=True)
        cdf.to_parquet(output_dir / f"event_{i}.parquet", index=False)
        event_specs.append({"name": f"event_{i}", "path": f"event_{i}.parquet", "source_table": cname,
                            "primary_key": cpk, "foreign_key": cfk, "time_column": ctime, "feature_columns": cfeats,
                            "rows": int(len(cdf))})
    meta = {
        "tables": [{"name": "root", "path": "root.parquet", "primary_key": root_pk, "time_column": root_time, "feature_columns": root_feats}],
        "event_tables": event_specs,
        "task": {"path": "task.parquet", "entity_key": entity_col, "time_column": time_col,
                 "target_column": target_col, "task_type": adapter_tt,
                 "num_classes": int(task_out[target_col].nunique()) if adapter_tt == "multiclass" else 2},
    }
    (output_dir / "metadata.json").write_text(json.dumps(json_safe(meta), allow_nan=False, indent=2, sort_keys=True) + "\n")
    return {"event_tables": [(e["source_table"], e["rows"]) for e in event_specs]}


def load_mt_model(ckpt_path, device):
    ck = torch.load(ckpt_path, map_location=device, weights_only=False)
    a = ck.get("manifest", {}).get("args", {})
    sd = ck["model"]
    out_dim = sd["readout.6.weight"].shape[0] if "readout.6.weight" in sd else sd["readout.3.weight"].shape[0]
    m = MultiTableRFM(
        root_cols=int(a.get("root_cols", 8)), event_cols=int(a.get("event_cols", 8)),
        num_event_tables=int(a.get("num_event_tables", 5)), lag_steps=int(a.get("lag_steps", 4)),
        max_classes=max(2, out_dim - 2), d_model=int(a.get("d_model", 256)), heads=int(a.get("heads", 8)),
        table_layers=int(a.get("table_layers", 2)), graph_layers=int(a.get("graph_layers", 3)),
        typed_task_conditioning=bool(a.get("typed_task_conditioning", False)) or ("task_cond.0.weight" in sd),
    ).to(device)
    m.load_state_dict(sd); m.eval()
    return m, int(a.get("num_event_tables", 5))


@torch.no_grad()
def evaluate_mt(model, dataset, num_event_tables, device, args):
    task_spec = dataset.specs["task"]; time_col = task_spec["time_column"]; tt = TASK_TYPES[task_spec["task_type"]]
    sorted_task = dataset.task.sort_values(time_col).reset_index(drop=True)
    train_end = temporal_train_end(sorted_task, time_col, train_frac=args.train_frac, context_size=args.context_size)
    if train_end <= 0 or train_end >= len(sorted_task):
        return {"error": "no split"}
    eval_rows = sorted_task.iloc[train_end:]
    num_eval = min(args.max_eval_samples, len(eval_rows))
    outs, ys, ycs, ncs = [], [], [], []
    nb = (num_eval + args.batch_size - 1) // args.batch_size
    for bi in range(nb):
        start = train_end + bi * args.batch_size; end = min(start + args.batch_size, train_end + num_eval)
        if start >= end: break
        b = dataset.to_batch(args.context_size, args.rows_per_event, args.lag_steps, args.root_cols,
                             args.event_cols, start, end - start, num_event_tables).to(device)
        o = model(b).float().cpu().numpy(); o = np.nan_to_num(o, nan=0.0, posinf=1e4, neginf=-1e4)
        outs.append(o); ys.append(b.y.cpu().numpy()); ycs.append(b.y_class.cpu().numpy()); ncs.append(b.num_classes.cpu().numpy())
    out = np.concatenate(outs); y = np.concatenate(ys)
    m = {"num_eval_samples": float(len(y)), "num_event_tables": float(num_event_tables)}
    if tt == 0 or (tt == 2 and int(ncs[0][0]) <= 2):
        m["binary_auroc"] = float(roc_auc_score(y, 1/(1+np.exp(-out[:, 0])))) if len(np.unique(y)) > 1 else float("nan")
    elif tt == 1:
        m["regression_mae"] = float(mean_absolute_error(y, out[:, 1]))
    return m


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--checkpoint", type=Path, required=True)
    p.add_argument("--dataset-dir", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--device", default="cuda")
    p.add_argument("--context-size", type=int, default=24)
    p.add_argument("--batch-size", type=int, default=8)
    p.add_argument("--rows-per-event", type=int, default=10)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--event-cols", type=int, default=8)
    p.add_argument("--lag-steps", type=int, default=4)
    p.add_argument("--max-eval-samples", type=int, default=512)
    p.add_argument("--train-frac", type=float, default=0.7)
    args = p.parse_args()
    device = torch.device(args.device if torch.cuda.is_available() or args.device == "cpu" else "cpu")
    model, net = load_mt_model(args.checkpoint, device)
    ds = MultiTableDataset.load(args.dataset_dir)
    metrics = evaluate_mt(model, ds, net, device, args)
    payload = {"checkpoint": str(args.checkpoint), "dataset_dir": str(args.dataset_dir), "metrics": metrics,
               "event_tables": [(e["source_table"], e["rows"]) for e in ds.specs["event_tables"]]}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    print(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
