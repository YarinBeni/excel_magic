"""Real-data export + loader for the N-table database-native RFM.

Captures the root entity table plus ALL FK->root event tables (up to a cap),
ranked by connectivity to the prediction entities, so star schemas (rel-stack,
rel-amazon) are represented faithfully instead of collapsed to a single child.
Builds MultiTableBatch with strict `< anchor` temporal filtering per event table
(no leakage), matching the eval-time ICL protocol.
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
import torch

from .multitable import MultiTableBatch
from .relational_io import (
    TASK_TYPES, temporal_train_end, table_values_with_time, numeric_values, lagged_targets, read_table,
)


class MultiTableDataset:
    def __init__(self, root: pd.DataFrame, events: list[pd.DataFrame], task: pd.DataFrame, specs: dict):
        self.root = root
        self.events = events
        self.task = task
        self.specs = specs

    @classmethod
    def load(cls, dataset_dir: Path | str) -> "MultiTableDataset":
        dataset_dir = Path(dataset_dir)
        specs = json.loads((dataset_dir / "metadata.json").read_text())
        root = read_table(dataset_dir / "root.parquet")
        events = [read_table(dataset_dir / e["path"]) for e in specs["event_tables"]]
        task = read_table(dataset_dir / specs["task"]["path"])
        return cls(root, events, task, specs)

    def to_batch(self, context_size, rows_per_event, lag_steps, root_cols, event_cols,
                 pred_start, batch_size, max_event_tables, context_mode="entity_mixed",
                 local_context_frac=0.5) -> MultiTableBatch:
        task_spec = self.specs["task"]
        entity_key = task_spec["entity_key"]; time_col = task_spec["time_column"]
        target_col = task_spec["target_column"]; task_type = TASK_TYPES[task_spec["task_type"]]
        num_classes = int(task_spec.get("num_classes", 2))
        task = self.task.sort_values(time_col).reset_index(drop=True)
        pred = task.iloc[pred_start: pred_start + batch_size]
        if len(pred) == 0:
            raise ValueError("no prediction rows")
        n = context_size + 1
        N = min(max_event_tables, len(self.events))
        root_spec = self.specs["tables"][0]
        root_pk = root_spec["primary_key"]
        root_feats = list(root_spec["feature_columns"])[:root_cols]
        ev_specs = self.specs["event_tables"][:N]
        max_time = max(float(task[time_col].max()), 1.0)
        root_indexed = self.root.drop_duplicates(root_pk, keep="last").set_index(root_pk)

        roots, events_all, ev_masks, tab_masks, targets, y_classes, times = [], [], [], [], [], [], []
        for _, prow in pred.iterrows():
            past = task[task[time_col] < prow[time_col]]
            if context_mode == "entity_mixed" and len(past):
                # Same-entity local history (so the model can read the entity's own past
                # labels to orient) mixed with global-recent context.
                n_local = int(round(context_size * local_context_frac))
                local = past[past[entity_key] == prow[entity_key]].tail(n_local)
                glob = past.tail(context_size - len(local))
                ctx = pd.concat([glob, local]).drop_duplicates().sort_values(time_col).tail(context_size)
            elif context_mode == "entity" and len(past):
                ctx = past[past[entity_key] == prow[entity_key]].tail(context_size)
            else:
                ctx = past.tail(context_size)
            ex = pd.concat([ctx, prow.to_frame().T], ignore_index=True)
            if len(ex) < n:
                pad = pd.DataFrame([{} for _ in range(n - len(ex))])
                ex = pd.concat([pad, ex], ignore_index=True)
            ex = ex.tail(n).reset_index(drop=True)
            root_a = np.zeros((n, root_cols), dtype=np.float32)
            events_a = np.zeros((n, N, rows_per_event, event_cols), dtype=np.float32)
            ev_mask = np.zeros((n, N, rows_per_event), dtype=bool)
            tab_mask = np.ones((n, N + 1), dtype=bool)
            target = np.zeros(n, dtype=np.float32); y_class = np.full(n, -1, np.int64)
            time_feat = np.zeros((n, 2), dtype=np.float32)
            for i, row in ex.iterrows():
                entity = row.get(entity_key); raw = row.get(time_col, 0.0)
                anchor = float(raw) if pd.notna(raw) else 0.0
                if entity in root_indexed.index and root_feats:
                    rv = root_indexed.loc[[entity], root_feats].head(1)
                    root_a[i, :len(root_feats)] = numeric_values(rv.iloc[0], root_feats)
                for t, espec in enumerate(ev_specs):
                    ev = self.events[t]; fk = espec["foreign_key"]; etime = espec.get("time_column")
                    if fk not in ev.columns:
                        tab_mask[i, t + 1] = False; continue
                    erows = ev[ev[fk] == entity]
                    if etime and etime in erows.columns:
                        erows = erows[erows[etime] < anchor]
                    erows = erows.tail(rows_per_event)
                    if len(erows):
                        feats = list(espec["feature_columns"])[:event_cols]
                        vals = table_values_with_time(erows, feats, etime, anchor, max_time, event_cols)
                        events_a[i, t, :len(vals), :vals.shape[1]] = vals
                        ev_mask[i, t, :len(vals)] = True
                if task_type == 2:
                    cls = int(row.get(target_col, 0)) if pd.notna(row.get(target_col, np.nan)) else 0
                    y_class[i] = max(0, min(num_classes - 1, cls)); target[i] = y_class[i] / max(1, num_classes - 1)
                else:
                    target[i] = float(row.get(target_col, 0.0)) if pd.notna(row.get(target_col, np.nan)) else 0.0
                time_feat[i, 0] = anchor / max_time
                time_feat[i, 1] = 1.0 if i >= max(0, n - max(1, n // 4)) else 0.0
            roots.append(root_a); events_all.append(events_a); ev_masks.append(ev_mask)
            tab_masks.append(tab_mask); targets.append(target); y_classes.append(y_class); times.append(time_feat)

        target_np = np.stack(targets).astype(np.float32)
        qm = np.zeros_like(target_np, dtype=bool); qm[:, -1] = True
        ctx_t = target_np.copy(); ctx_t[:, -1] = 0.0
        lag = lagged_targets(target_np, np.stack(times)[:, :, 0], lag_steps)
        return MultiTableBatch(
            root=torch.tensor(np.stack(roots), dtype=torch.float32),
            events=torch.tensor(np.stack(events_all), dtype=torch.float32),
            event_mask=torch.tensor(np.stack(ev_masks), dtype=torch.bool),
            table_mask=torch.tensor(np.stack(tab_masks), dtype=torch.bool),
            context_target=torch.tensor(ctx_t, dtype=torch.float32),
            lag_target=torch.tensor(lag, dtype=torch.float32),
            time=torch.tensor(np.stack(times), dtype=torch.float32),
            query_mask=torch.tensor(qm, dtype=torch.bool),
            task_type=torch.full((len(pred),), task_type, dtype=torch.long),
            num_classes=torch.full((len(pred),), num_classes, dtype=torch.long),
            y=torch.tensor(target_np[:, -1], dtype=torch.float32),
            y_class=torch.tensor(np.stack(y_classes)[:, -1], dtype=torch.long),
        )
