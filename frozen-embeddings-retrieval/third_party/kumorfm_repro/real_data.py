"""Real-database co-training sampler for KumoRFM-style pre-training.

KumoRFM-2 is pre-trained on an expanded mix of synthetic SCM data and real
relational databases. This module draws in-context tasks from exported RelBench
datasets so the same model can be co-trained on real relational structure.

Leakage control:
  * Used under leave-one-dataset-out (LODO): the evaluation dataset family is
    never passed in ``dataset_dirs``.
  * Each in-context batch is materialized via ``RelationalDataset.to_batch`` with
    strict ``< anchor`` temporal filtering for child/aux/root rows and a hidden
    query label, identical to the eval-time protocol.
  * Anchors are sampled only from each dataset's training region (the first
    ``train_frac`` of unique timestamps), so future rows never enter context.
"""
from __future__ import annotations

from pathlib import Path

import numpy as np

from .relational_io import RelationalDataset, TASK_TYPES, temporal_train_end


class RealTaskSampler:
    def __init__(
        self,
        dataset_dirs: list[Path],
        *,
        context_size: int,
        rows_per_child: int,
        rows_per_aux: int,
        lag_steps: int,
        root_cols: int,
        child_cols: int,
        aux_cols: int,
        train_frac: float = 0.7,
        context_mode: str = "entity_mixed",
        local_context_frac: float = 0.5,
        seed: int = 0,
    ) -> None:
        self.context_size = context_size
        self.rows_per_child = rows_per_child
        self.rows_per_aux = rows_per_aux
        self.lag_steps = lag_steps
        self.root_cols = root_cols
        self.child_cols = child_cols
        self.aux_cols = aux_cols
        self.context_mode = context_mode
        self.local_context_frac = local_context_frac
        self.rng = np.random.default_rng(seed)

        self.datasets: list[RelationalDataset] = []
        self.train_ends: list[int] = []
        self.time_cols: list[str] = []
        for d in dataset_dirs:
            d = Path(d)
            if not (d / "metadata.json").exists():
                continue
            ds = RelationalDataset.load(d)
            time_col = ds.specs["task"]["time_column"]
            sorted_task = ds.task.sort_values(time_col).reset_index(drop=True)
            ds.task = sorted_task
            train_end = temporal_train_end(sorted_task, time_col, train_frac=train_frac, context_size=context_size)
            # Need room for at least a few query anchors strictly after some context.
            if train_end <= context_size + 1:
                continue
            self.datasets.append(ds)
            self.train_ends.append(int(train_end))
            self.time_cols.append(time_col)
        if not self.datasets:
            raise ValueError("RealTaskSampler: no usable datasets after filtering.")

    def __len__(self) -> int:
        return len(self.datasets)

    @property
    def names(self) -> list[str]:
        return [ds.specs.get("source", {}).get("root_table", "?") for ds in self.datasets]

    def batch(self, batch_size: int):
        # Pick one dataset per batch so all queries share a schema (model handles
        # varying column counts via positional column embeddings).
        idx = int(self.rng.integers(0, len(self.datasets)))
        ds = self.datasets[idx]
        train_end = self.train_ends[idx]
        low = self.context_size + 1
        high = max(low + 1, train_end - batch_size)
        pred_start = int(self.rng.integers(low, high)) if high > low else low
        batch = ds.to_batch(
            context_size=self.context_size,
            rows_per_child=self.rows_per_child,
            rows_per_aux=self.rows_per_aux,
            lag_steps=self.lag_steps,
            root_cols=self.root_cols,
            child_cols=self.child_cols,
            aux_cols=self.aux_cols,
            pred_start=pred_start,
            batch_size=batch_size,
            context_mode=self.context_mode,
            local_context_frac=self.local_context_frac,
        )
        return batch
