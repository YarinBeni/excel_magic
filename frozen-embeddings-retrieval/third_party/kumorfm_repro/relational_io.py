from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import torch
from sklearn.ensemble import HistGradientBoostingClassifier, HistGradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, roc_auc_score

from .data import RelationalBatch


TASK_TYPES = {"binary": 0, "regression": 1, "multiclass": 2}


def temporal_train_end(task: pd.DataFrame, time_col: str, train_frac: float = 0.7, context_size: int = 0) -> int:
    """Return a row split whose eval rows start at a strictly later timestamp."""
    sorted_task = task.sort_values(time_col).reset_index(drop=True)
    if len(sorted_task) == 0:
        return 0
    min_train_rows = max(1, int(context_size) + 1)
    unique_times = np.array(sorted_task[time_col].dropna().unique())
    unique_times.sort()
    if len(unique_times) < 2:
        return min(len(sorted_task), max(min_train_rows, int(len(sorted_task) * train_frac)))

    split_idx = int(len(unique_times) * train_frac)
    split_idx = max(1, min(split_idx, len(unique_times) - 1))
    while split_idx < len(unique_times) - 1 and int((sorted_task[time_col] < unique_times[split_idx]).sum()) < min_train_rows:
        split_idx += 1
    return int((sorted_task[time_col] < unique_times[split_idx]).sum())


@dataclass(frozen=True)
class TableSpec:
    name: str
    path: str
    primary_key: str
    time_column: str | None = None
    foreign_key: str | None = None
    parent_table: str | None = None
    feature_columns: tuple[str, ...] = ()


@dataclass(frozen=True)
class TaskSpec:
    entity_table: str
    entity_key: str
    time_column: str
    target_column: str
    task_type: str
    num_classes: int = 2


class RelationalDataset:
    """Small file-backed relational dataset adapter.

    This is intentionally conservative: it supports a root/entity table, one
    child table, and one auxiliary table, matching the current model surface.
    It is designed so benchmark exports such as RelBench task tables can be
    converted without changing model code.
    """

    def __init__(self, root: pd.DataFrame, child: pd.DataFrame, aux: pd.DataFrame, task: pd.DataFrame, specs: dict[str, Any]):
        self.root = root
        self.child = child
        self.aux = aux
        self.task = task
        self.specs = specs
        self._index_cache: dict[str, Any] = {}

    @classmethod
    def load(cls, dataset_dir: Path | str) -> "RelationalDataset":
        dataset_dir = Path(dataset_dir)
        specs = json.loads((dataset_dir / "metadata.json").read_text())
        tables = {}
        for table in specs["tables"]:
            tables[table["name"]] = read_table(dataset_dir / table["path"])
        task = read_table(dataset_dir / specs["task"]["path"])
        return cls(
            root=tables[specs["root_table"]],
            child=tables[specs["child_table"]],
            aux=tables[specs["aux_table"]],
            task=task,
            specs=specs,
        )

    def to_batch(
        self,
        context_size: int,
        rows_per_child: int,
        rows_per_aux: int,
        lag_steps: int,
        root_cols: int,
        child_cols: int,
        aux_cols: int,
        pred_start: int | None = None,
        batch_size: int = 16,
        context_mode: str = "recent",
        local_context_frac: float = 0.5,
        indexed_table_lookup: bool = False,
    ) -> RelationalBatch:
        task_spec = self.specs["task"]
        entity_key = task_spec["entity_key"]
        time_col = task_spec["time_column"]
        target_col = task_spec["target_column"]
        task_type = TASK_TYPES[task_spec["task_type"]]

        task = self._sorted_task(time_col)
        task_times = self._task_time_values(time_col)
        task_entity_indices = self._task_entity_indices(entity_key, time_col)
        if pred_start is None:
            pred_start = max(context_size, len(task) - batch_size)
        pred = task.iloc[pred_start : pred_start + batch_size]
        if len(pred) == 0:
            raise ValueError("No prediction rows selected")

        root_rows = []
        child_rows = []
        aux_rows = []
        child_masks = []
        aux_masks = []
        table_masks = []
        targets = []
        y_classes = []
        times = []

        for _, pred_row in pred.iterrows():
            context, context_local = self._select_context_examples(
                task,
                task_times,
                task_entity_indices,
                pred_row,
                entity_key,
                time_col,
                context_size,
                context_mode=context_mode,
                local_context_frac=local_context_frac,
            )
            examples = pd.concat([context, pred_row.to_frame().T], ignore_index=True)
            local_flags = np.concatenate([context_local, np.array([1.0], dtype=np.float32)])
            examples = pad_examples(examples, context_size + 1)
            local_flags = pad_local_flags(local_flags, context_size + 1)
            root_arr, child_arr, aux_arr, child_mask, aux_mask, table_mask, target, y_class, time_feat = self._examples_to_arrays(
                examples,
                entity_key,
                time_col,
                target_col,
                rows_per_child,
                rows_per_aux,
                root_cols,
                child_cols,
                aux_cols,
                task_type,
                task_spec.get("num_classes", 2),
                local_flags,
                indexed_table_lookup,
            )
            root_rows.append(root_arr)
            child_rows.append(child_arr)
            aux_rows.append(aux_arr)
            child_masks.append(child_mask)
            aux_masks.append(aux_mask)
            table_masks.append(table_mask)
            targets.append(target)
            y_classes.append(y_class)
            times.append(time_feat)

        target_np = np.stack(targets).astype(np.float32)
        query_mask = np.zeros_like(target_np, dtype=bool)
        query_mask[:, -1] = True
        context_target = target_np.copy()
        context_target[:, -1] = 0.0
        lag_target = lagged_targets(target_np, np.stack(times)[:, :, 0], lag_steps)
        task_type_arr = np.full(len(pred), task_type, dtype=np.int64)
        num_classes_arr = np.full(len(pred), task_spec.get("num_classes", 2), dtype=np.int64)

        return RelationalBatch(
            root=torch.tensor(np.stack(root_rows), dtype=torch.float32),
            child=torch.tensor(np.stack(child_rows), dtype=torch.float32),
            aux=torch.tensor(np.stack(aux_rows), dtype=torch.float32),
            child_mask=torch.tensor(np.stack(child_masks), dtype=torch.bool),
            aux_mask=torch.tensor(np.stack(aux_masks), dtype=torch.bool),
            table_mask=torch.tensor(np.stack(table_masks), dtype=torch.bool),
            context_target=torch.tensor(context_target, dtype=torch.float32),
            lag_target=torch.tensor(lag_target, dtype=torch.float32),
            time=torch.tensor(np.stack(times), dtype=torch.float32),
            query_mask=torch.tensor(query_mask, dtype=torch.bool),
            task_type=torch.tensor(task_type_arr, dtype=torch.long),
            mechanism=torch.full((len(pred),), -1, dtype=torch.long),
            num_classes=torch.tensor(num_classes_arr, dtype=torch.long),
            y=torch.tensor(target_np[:, -1], dtype=torch.float32),
            y_class=torch.tensor(np.stack(y_classes)[:, -1], dtype=torch.long),
            target_all=torch.tensor(target_np, dtype=torch.float32),
            y_class_all=torch.tensor(np.stack(y_classes), dtype=torch.long),
        )

    def _sorted_task(self, time_col: str) -> pd.DataFrame:
        key = f"task:{time_col}"
        if key not in self._index_cache:
            self._index_cache[key] = self.task.sort_values(time_col).reset_index(drop=True)
        return self._index_cache[key]

    def _task_time_values(self, time_col: str) -> np.ndarray:
        key = f"task_times:{time_col}"
        if key not in self._index_cache:
            task = self._sorted_task(time_col)
            self._index_cache[key] = pd.to_numeric(task[time_col], errors="coerce").fillna(0.0).to_numpy(dtype=np.float64)
        return self._index_cache[key]

    def _task_entity_indices(self, entity_key: str, time_col: str) -> dict[Any, np.ndarray]:
        key = f"task_entities:{time_col}:{entity_key}"
        if key not in self._index_cache:
            task = self._sorted_task(time_col)
            groups: dict[Any, np.ndarray] = {}
            for entity, frame in task.groupby(entity_key, sort=False, dropna=False):
                groups[entity] = frame.index.to_numpy(dtype=np.int64)
            self._index_cache[key] = groups
        return self._index_cache[key]

    def _select_context_examples(
        self,
        task: pd.DataFrame,
        task_times: np.ndarray,
        task_entity_indices: dict[Any, np.ndarray],
        pred_row: pd.Series,
        entity_key: str,
        time_col: str,
        context_size: int,
        context_mode: str = "recent",
        local_context_frac: float = 0.5,
    ) -> tuple[pd.DataFrame, np.ndarray]:
        if context_size <= 0:
            return task.iloc[0:0], np.zeros(0, dtype=np.float32)
        anchor = float(pred_row[time_col]) if pd.notna(pred_row[time_col]) else 0.0
        history_end = int(np.searchsorted(task_times, anchor, side="left"))
        if history_end <= 0:
            return task.iloc[0:0], np.zeros(0, dtype=np.float32)
        if context_mode == "recent":
            start = max(0, history_end - context_size)
            selected_idx = np.arange(start, history_end, dtype=np.int64)
            return task.iloc[selected_idx], np.zeros(len(selected_idx), dtype=np.float32)
        if context_mode not in {"entity", "entity_mixed"}:
            raise ValueError(f"Unknown context_mode: {context_mode}")

        entity_indices = task_entity_indices.get(pred_row[entity_key], np.zeros(0, dtype=np.int64))
        local_history = entity_indices[entity_indices < history_end]
        local_target = context_size if context_mode == "entity" else int(np.ceil(context_size * float(np.clip(local_context_frac, 0.0, 1.0))))
        local_idx = local_history[-local_target:] if local_target > 0 else np.zeros(0, dtype=np.int64)

        remaining = max(0, context_size - len(local_idx))
        history_idx = np.arange(max(0, history_end - context_size - len(local_idx)), history_end, dtype=np.int64)
        if len(local_idx):
            global_pool = history_idx[~np.isin(history_idx, local_idx, assume_unique=False)]
        else:
            global_pool = history_idx
        global_idx = global_pool[-remaining:] if remaining > 0 else np.zeros(0, dtype=np.int64)

        selected_idx = np.concatenate([global_idx, local_idx])
        flags = np.concatenate([np.zeros(len(global_idx), dtype=np.float32), np.ones(len(local_idx), dtype=np.float32)])
        if len(selected_idx) < context_size:
            topup_needed = context_size - len(selected_idx)
            wider_history = np.arange(0, history_end, dtype=np.int64)
            if len(selected_idx):
                topup_pool = wider_history[~np.isin(wider_history, selected_idx, assume_unique=False)]
            else:
                topup_pool = wider_history
            topup_idx = topup_pool[-topup_needed:]
            selected_idx = np.concatenate([topup_idx, selected_idx])
            flags = np.concatenate([np.zeros(len(topup_idx), dtype=np.float32), flags])
        if len(selected_idx) > context_size:
            selected_idx = selected_idx[-context_size:]
            flags = flags[-context_size:]
        return task.iloc[selected_idx].reset_index(drop=True), flags.astype(np.float32, copy=False)

    def _grouped_table(self, table_name: str, frame: pd.DataFrame, key_col: str, time_col: str | None = None) -> dict[Any, pd.DataFrame]:
        cache_key = f"group:{table_name}:{key_col}:{time_col or ''}"
        if cache_key not in self._index_cache:
            table = frame
            if time_col and time_col in table.columns:
                table = table.sort_values(time_col)
            self._index_cache[cache_key] = {key: group for key, group in table.groupby(key_col, sort=False, dropna=False)}
        return self._index_cache[cache_key]

    def validate(self, train_frac: float = 0.7, context_size: int = 16) -> dict[str, Any]:
        errors: list[str] = []
        warnings: list[str] = []
        task_spec = self.specs["task"]
        table_specs = {table["name"]: table for table in self.specs["tables"]}
        for required in ("root_table", "child_table", "aux_table", "task"):
            if required not in self.specs:
                errors.append(f"metadata missing {required}")
        for table_name, frame in ((self.specs.get("root_table"), self.root), (self.specs.get("child_table"), self.child), (self.specs.get("aux_table"), self.aux)):
            if table_name is None:
                continue
            spec = table_specs.get(table_name, {})
            for column in [spec.get("primary_key"), spec.get("foreign_key"), spec.get("time_column"), *spec.get("feature_columns", [])]:
                if column and column not in frame.columns:
                    errors.append(f"table {table_name} missing column {column}")

        for column in (task_spec.get("entity_key"), task_spec.get("time_column"), task_spec.get("target_column")):
            if column not in self.task.columns:
                errors.append(f"task table missing column {column}")

        time_col = task_spec["time_column"]
        sorted_task = self.task.sort_values(time_col).reset_index(drop=True)
        train_end = temporal_train_end(sorted_task, time_col, train_frac=train_frac, context_size=context_size)
        split = {
            "rows_total": int(len(sorted_task)),
            "rows_train": int(train_end),
            "rows_eval": int(max(0, len(sorted_task) - train_end)),
            "train_max_time": float(sorted_task.iloc[:train_end][time_col].max()) if train_end > 0 else None,
            "eval_min_time": float(sorted_task.iloc[train_end:][time_col].min()) if train_end < len(sorted_task) else None,
        }
        if split["rows_eval"] <= 0:
            errors.append("temporal split has no evaluation rows")
        if split["train_max_time"] is not None and split["eval_min_time"] is not None and split["train_max_time"] >= split["eval_min_time"]:
            errors.append("temporal split is not strictly ordered by time")

        context_counts = []
        future_child_rows = 0
        future_aux_rows = 0
        child_spec = table_specs.get(self.specs.get("child_table"), {})
        aux_spec = table_specs.get(self.specs.get("aux_table"), {})
        entity_key = task_spec["entity_key"]
        for _, row in sorted_task.iloc[train_end:].head(128).iterrows():
            anchor = row[time_col]
            context_counts.append(int((sorted_task[time_col] < anchor).sum()))
            entity = row[entity_key]
            if child_spec.get("time_column") and child_spec.get("foreign_key") in self.child.columns:
                future_child_rows += int(((self.child[child_spec["foreign_key"]] == entity) & (self.child[child_spec["time_column"]] >= anchor)).sum())
            if aux_spec.get("time_column") and child_spec.get("primary_key") and aux_spec.get("foreign_key"):
                child_rows = self.child[self.child[child_spec["foreign_key"]] == entity] if child_spec.get("foreign_key") in self.child.columns else self.child.iloc[0:0]
                if child_spec.get("time_column") and child_spec["time_column"] in child_rows.columns:
                    child_rows = child_rows[child_rows[child_spec["time_column"]] < anchor]
                aux_rows = self.aux[self.aux[aux_spec["foreign_key"]].isin(child_rows[child_spec["primary_key"]])] if child_spec.get("primary_key") in child_rows.columns and aux_spec.get("foreign_key") in self.aux.columns else self.aux.iloc[0:0]
                future_aux_rows += int((aux_rows[aux_spec["time_column"]] >= anchor).sum()) if aux_spec["time_column"] in aux_rows.columns else 0
        context = {
            "min_available_context": int(min(context_counts)) if context_counts else 0,
            "median_available_context": float(np.median(context_counts)) if context_counts else 0.0,
            "requested_context_size": int(context_size),
        }
        temporal_filtering = {
            "future_child_rows_excluded_in_eval_sample": int(future_child_rows),
            "future_aux_rows_excluded_in_eval_sample": int(future_aux_rows),
            "checked_eval_rows": int(min(128, split["rows_eval"])),
        }
        if context["min_available_context"] < context_size:
            warnings.append("some evaluation rows have fewer historical context examples than requested")
        return {
            "ok": len(errors) == 0,
            "errors": errors,
            "warnings": warnings,
            "split": split,
            "context": context,
            "temporal_filtering": temporal_filtering,
            "task_type": task_spec.get("task_type"),
        }

    def _examples_to_arrays(
        self,
        examples: pd.DataFrame,
        entity_key: str,
        time_col: str,
        target_col: str,
        rows_per_child: int,
        rows_per_aux: int,
        root_cols: int,
        child_cols: int,
        aux_cols: int,
        task_type: int,
        num_classes: int,
        local_flags: np.ndarray | None = None,
        indexed_table_lookup: bool = False,
    ):
        n = len(examples)
        root = np.zeros((n, root_cols), dtype=np.float32)
        child = np.zeros((n, rows_per_child, child_cols), dtype=np.float32)
        aux = np.zeros((n, rows_per_aux, aux_cols), dtype=np.float32)
        child_mask = np.zeros((n, rows_per_child), dtype=bool)
        aux_mask = np.zeros((n, rows_per_aux), dtype=bool)
        table_mask = np.ones((n, 3), dtype=bool)
        target = np.zeros(n, dtype=np.float32)
        y_class = np.full(n, -1, dtype=np.int64)
        time_feat = np.zeros((n, 2), dtype=np.float32)

        root_spec, child_spec, aux_spec = self.specs["tables"]
        root_features = list(root_spec["feature_columns"])[:root_cols]
        child_features = list(child_spec["feature_columns"])[:child_cols]
        aux_features = list(aux_spec["feature_columns"])[:aux_cols]
        max_time = max(float(self.task[time_col].max()), 1.0)
        root_groups = child_groups = aux_groups = None
        if indexed_table_lookup:
            root_groups = self._grouped_table(root_spec["name"], self.root, root_spec["primary_key"], root_spec.get("time_column"))
            child_groups = self._grouped_table(child_spec["name"], self.child, child_spec["foreign_key"], child_spec.get("time_column"))
            aux_fk = aux_spec.get("foreign_key")
            if aux_fk:
                aux_groups = self._grouped_table(aux_spec["name"], self.aux, aux_fk, aux_spec.get("time_column"))

        for i, row in examples.iterrows():
            entity = row.get(entity_key)
            raw_anchor = row.get(time_col, 0.0)
            anchor = float(raw_anchor) if pd.notna(raw_anchor) else 0.0
            if root_groups is not None:
                root_candidates = root_groups.get(entity, self.root.iloc[0:0])
            else:
                root_candidates = self.root[self.root[root_spec["primary_key"]] == entity]
            if root_spec.get("time_column") and root_spec["time_column"] in root_candidates.columns:
                root_candidates = root_candidates[root_candidates[root_spec["time_column"]] < anchor]
                if root_groups is None:
                    root_candidates = root_candidates.sort_values(root_spec["time_column"])
            root_match = root_candidates.tail(1)
            if len(root_match):
                root[i, : len(root_features)] = numeric_values(root_match.iloc[0], root_features)

            if child_groups is not None:
                child_rows = child_groups.get(entity, self.child.iloc[0:0])
            else:
                child_rows = self.child[self.child[child_spec["foreign_key"]] == entity]
            if child_spec.get("time_column"):
                child_rows = child_rows[child_rows[child_spec["time_column"]] < anchor]
            child_rows = child_rows.tail(rows_per_child)
            if len(child_rows):
                values = table_values_with_time(
                    child_rows,
                    child_features,
                    child_spec.get("time_column"),
                    anchor,
                    max_time,
                    child_cols,
                )
                child[i, : len(values), : values.shape[1]] = values
                child_mask[i, : len(values)] = True

            if aux_spec.get("parent_table") == "root" and aux_spec.get("foreign_key") in self.aux.columns:
                # Star schema: aux is a second sibling child keyed directly to the root entity.
                aux_rows = self.aux[self.aux[aux_spec["foreign_key"]] == entity]
            else:
                aux_key = child_spec.get("primary_key")
                if aux_key and aux_groups is not None and len(child_rows):
                    aux_parts = [aux_groups.get(key) for key in child_rows[aux_key]]
                    aux_rows = pd.concat([part for part in aux_parts if part is not None], axis=0) if any(part is not None for part in aux_parts) else self.aux.iloc[0:0]
                elif aux_key and aux_spec.get("foreign_key") and len(child_rows):
                    aux_rows = self.aux[self.aux[aux_spec["foreign_key"]].isin(child_rows[aux_key])]
                else:
                    aux_rows = self.aux.iloc[0:0]
            if aux_spec.get("time_column"):
                aux_rows = aux_rows[aux_rows[aux_spec["time_column"]] < anchor]
            aux_rows = aux_rows.tail(rows_per_aux)
            if len(aux_rows):
                values = table_values_with_time(
                    aux_rows,
                    aux_features,
                    aux_spec.get("time_column"),
                    anchor,
                    max_time,
                    aux_cols,
                )
                aux[i, : len(values), : values.shape[1]] = values
                aux_mask[i, : len(values)] = True

            if task_type == 2:
                cls = int(row.get(target_col, 0)) if pd.notna(row.get(target_col, np.nan)) else 0
                y_class[i] = max(0, min(num_classes - 1, cls))
                target[i] = y_class[i] / max(1, num_classes - 1)
            else:
                target[i] = float(row.get(target_col, 0.0)) if pd.notna(row.get(target_col, np.nan)) else 0.0
            time_feat[i, 0] = anchor / max_time
            if local_flags is not None and i < len(local_flags):
                time_feat[i, 1] = float(local_flags[i])
            else:
                time_feat[i, 1] = 1.0 if i >= max(0, n - max(1, n // 4)) else 0.0
        return root, child, aux, child_mask, aux_mask, table_mask, target, y_class, time_feat


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix == ".parquet":
        return pd.read_parquet(path)
    if path.suffix == ".csv":
        return pd.read_csv(path)
    raise ValueError(f"Unsupported table format: {path}")


def write_demo_dataset(dataset_dir: Path | str, rows: int = 256, seed: int = 0) -> None:
    dataset_dir = Path(dataset_dir)
    dataset_dir.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(seed)
    entities = np.arange(rows)
    root = pd.DataFrame({"entity_id": entities, "root_x0": rng.normal(size=rows), "root_x1": rng.normal(size=rows)})
    child = pd.DataFrame(
        {
            "event_id": np.arange(rows * 3),
            "entity_id": np.repeat(entities, 3),
            "time": np.tile(np.arange(3), rows) + np.repeat(entities, 3) * 0.01,
            "child_x0": rng.normal(size=rows * 3),
            "child_x1": rng.normal(size=rows * 3),
        }
    )
    aux = pd.DataFrame(
        {
            "aux_id": np.arange(rows * 3),
            "event_id": np.arange(rows * 3),
            "time": child["time"].to_numpy() + 0.1,
            "aux_x0": rng.normal(size=rows * 3),
            "aux_x1": rng.normal(size=rows * 3),
        }
    )
    child_signal = child.groupby("entity_id")["child_x0"].mean()
    aux_signal = aux.join(child[["event_id", "entity_id"]].set_index("event_id"), on="event_id").groupby("entity_id")["aux_x0"].mean()
    score = root.set_index("entity_id")["root_x0"] + child_signal + aux_signal
    target = (score > score.median()).astype(int)
    task = pd.DataFrame({"entity_id": entities, "time": entities.astype(float), "target": target.to_numpy()})

    root.to_parquet(dataset_dir / "root.parquet")
    child.to_parquet(dataset_dir / "child.parquet")
    aux.to_parquet(dataset_dir / "aux.parquet")
    task.to_parquet(dataset_dir / "task.parquet")
    metadata = {
        "root_table": "root",
        "child_table": "child",
        "aux_table": "aux",
        "tables": [
            {"name": "root", "path": "root.parquet", "primary_key": "entity_id", "feature_columns": ["root_x0", "root_x1"]},
            {
                "name": "child",
                "path": "child.parquet",
                "primary_key": "event_id",
                "foreign_key": "entity_id",
                "parent_table": "root",
                "time_column": "time",
                "feature_columns": ["child_x0", "child_x1"],
            },
            {
                "name": "aux",
                "path": "aux.parquet",
                "primary_key": "aux_id",
                "foreign_key": "event_id",
                "parent_table": "child",
                "time_column": "time",
                "feature_columns": ["aux_x0", "aux_x1"],
            },
        ],
        "task": {
            "path": "task.parquet",
            "entity_table": "root",
            "entity_key": "entity_id",
            "time_column": "time",
            "target_column": "target",
            "task_type": "binary",
            "num_classes": 2,
        },
    }
    (dataset_dir / "metadata.json").write_text(json.dumps(metadata, indent=2, sort_keys=True) + "\n")


def evaluate_flattened_baseline(dataset: RelationalDataset, train_frac: float = 0.7) -> dict[str, float]:
    frame = flatten_dataset(dataset)
    task = dataset.specs["task"]
    target = task["target_column"]
    train_n = temporal_train_end(frame, task["time_column"], train_frac=train_frac)
    train = frame.iloc[:train_n]
    test = frame.iloc[train_n:]
    feature_cols = [c for c in frame.columns if c not in {target, task["entity_key"], task["time_column"]}]
    if len(train) == 0 or len(test) == 0:
        raise ValueError("Not enough rows for a train/test split")
    if task["task_type"] == "regression":
        model = HistGradientBoostingRegressor(random_state=0)
        model.fit(train[feature_cols], train[target])
        pred = model.predict(test[feature_cols])
        return {"regression_mae": float(mean_absolute_error(test[target], pred))}
    model = HistGradientBoostingClassifier(random_state=0)
    model.fit(train[feature_cols], train[target])
    proba = model.predict_proba(test[feature_cols])
    if task["task_type"] == "binary":
        score = proba[:, 1] if proba.shape[1] > 1 else proba[:, 0]
        return {"binary_auroc": float(roc_auc_score(test[target], score))}
    class_to_col = {int(label): idx for idx, label in enumerate(model.classes_)}
    fallback_rank = len(model.classes_) + 1
    ranks = []
    for row, label in zip(proba, test[target]):
        label = int(label)
        if label not in class_to_col:
            ranks.append(1.0 / fallback_rank)
            continue
        order = np.argsort(-row)
        col = class_to_col[label]
        ranks.append(1.0 / (int(np.where(order == col)[0][0]) + 1))
    return {"multiclass_mrr": float(np.mean(ranks))}


def flatten_dataset(dataset: RelationalDataset) -> pd.DataFrame:
    task = dataset.task.copy().sort_values(dataset.specs["task"]["time_column"]).reset_index(drop=True)
    root_spec, child_spec, aux_spec = dataset.specs["tables"]
    root_features = list(root_spec["feature_columns"])
    child_features = list(child_spec["feature_columns"])
    aux_features = list(aux_spec["feature_columns"])
    root = dataset.root[[root_spec["primary_key"], *root_features]]
    if root_spec.get("time_column"):
        root = root.sort_values(root_spec["time_column"])
    child = dataset.child.copy()
    if child_spec.get("time_column"):
        child = child.sort_values(child_spec["time_column"])
    aux = dataset.aux.copy()
    if aux_spec.get("time_column"):
        aux = aux.sort_values(aux_spec["time_column"])
    rows = []
    for _, task_row in task.iterrows():
        anchor = float(task_row[dataset.specs["task"]["time_column"]])
        entity = task_row[dataset.specs["task"]["entity_key"]]
        root_sub = root[root[root_spec["primary_key"]] == entity]
        if root_spec.get("time_column") and root_spec["time_column"] in root_sub.columns:
            root_sub = root_sub[root_sub[root_spec["time_column"]] < anchor]
        root_vals = root_sub[root_features].tail(1)
        child_sub = child[child[child_spec["foreign_key"]] == entity]
        if child_spec.get("time_column"):
            child_sub = child_sub[child_sub[child_spec["time_column"]] < anchor]
        child_agg = child_sub[child_features].agg(["mean", "max", "min"]).unstack().to_frame().T if len(child_sub) else pd.DataFrame(columns=[f"child_{f}_{a}" for f in child_features for a in ["mean", "max", "min"]])
        child_agg.columns = [f"child_{a}_{b}" for a, b in child_agg.columns] if not child_agg.empty else child_agg.columns
        aux_vals = pd.DataFrame()
        if len(child_sub) and aux_spec.get("foreign_key"):
            child_keys = child_sub[child_spec["primary_key"]]
            aux_sub = aux[aux[aux_spec["foreign_key"]].isin(child_keys)]
            if aux_spec.get("time_column"):
                aux_sub = aux_sub[aux_sub[aux_spec["time_column"]] < anchor]
            if len(aux_sub):
                aux_agg = aux_sub[aux_features].agg(["mean", "max", "min"]).unstack().to_frame().T
                aux_agg.columns = [f"aux_{a}_{b}" for a, b in aux_agg.columns]
                aux_vals = aux_agg
        row_data = pd.concat([root_vals.reset_index(drop=True), child_agg.reset_index(drop=True), aux_vals.reset_index(drop=True)], axis=1)
        row_data[dataset.specs["task"]["entity_key"]] = entity
        row_data[dataset.specs["task"]["time_column"]] = anchor
        row_data[dataset.specs["task"]["target_column"]] = task_row[dataset.specs["task"]["target_column"]]
        rows.append(row_data)
    if not rows:
        return pd.DataFrame()
    frame = pd.concat(rows, ignore_index=True)
    return frame.fillna(0.0)


def pad_examples(examples: pd.DataFrame, n: int) -> pd.DataFrame:
    if len(examples) >= n:
        return examples.tail(n).reset_index(drop=True)
    padding = pd.DataFrame([{} for _ in range(n - len(examples))])
    return pd.concat([padding, examples], ignore_index=True).reset_index(drop=True)


def pad_local_flags(flags: np.ndarray, n: int) -> np.ndarray:
    flags = np.asarray(flags, dtype=np.float32)
    if len(flags) >= n:
        return flags[-n:]
    padding = np.zeros(n - len(flags), dtype=np.float32)
    return np.concatenate([padding, flags]).astype(np.float32, copy=False)


def select_context_examples(
    task: pd.DataFrame,
    pred_row: pd.Series,
    entity_key: str,
    time_col: str,
    context_size: int,
    context_mode: str = "recent",
    local_context_frac: float = 0.5,
) -> tuple[pd.DataFrame, np.ndarray]:
    if context_size <= 0:
        return task.iloc[0:0], np.zeros(0, dtype=np.float32)
    anchor = pred_row[time_col]
    history = task[task[time_col] < anchor]
    if context_mode == "recent" or len(history) == 0:
        context = history.tail(context_size)
        return context, np.zeros(len(context), dtype=np.float32)
    if context_mode not in {"entity", "entity_mixed"}:
        raise ValueError(f"Unknown context_mode: {context_mode}")

    local_history = history[history[entity_key] == pred_row[entity_key]]
    if context_mode == "entity":
        local_target = context_size
    else:
        frac = float(np.clip(local_context_frac, 0.0, 1.0))
        local_target = int(np.ceil(context_size * frac))
    local = local_history.tail(local_target)

    remaining = max(0, context_size - len(local))
    global_pool = history.drop(index=local.index, errors="ignore")
    global_context = global_pool.tail(remaining)

    selected = pd.concat([global_context, local], axis=0)
    if len(selected) < context_size:
        topup_pool = history.drop(index=selected.index, errors="ignore")
        topup = topup_pool.tail(context_size - len(selected))
        selected = pd.concat([topup, selected], axis=0)
        flags = np.concatenate(
            [
                np.zeros(len(topup) + len(global_context), dtype=np.float32),
                np.ones(len(local), dtype=np.float32),
            ]
        )
    else:
        flags = np.concatenate([np.zeros(len(global_context), dtype=np.float32), np.ones(len(local), dtype=np.float32)])
    if len(selected) > context_size:
        selected = selected.tail(context_size)
        flags = flags[-context_size:]
    return selected.reset_index(drop=True), flags.astype(np.float32, copy=False)


def numeric_values(row: pd.Series, columns: list[str]) -> np.ndarray:
    return pd.to_numeric(row[columns], errors="coerce").fillna(0.0).to_numpy(dtype=np.float32)


def table_values_with_time(
    rows: pd.DataFrame,
    feature_columns: list[str],
    time_column: str | None,
    anchor: float,
    max_time: float,
    width: int,
) -> np.ndarray:
    base_width = min(len(feature_columns), width)
    if base_width:
        base = rows[feature_columns[:base_width]].apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=np.float32)
    else:
        base = np.zeros((len(rows), 0), dtype=np.float32)
    remaining = max(0, width - base.shape[1])
    if remaining and time_column and time_column in rows.columns:
        row_time = pd.to_numeric(rows[time_column], errors="coerce").fillna(0.0).to_numpy(dtype=np.float64)
        denom = max(float(max_time), 1.0)
        rel_age = np.maximum(float(anchor) - row_time, 0.0)
        temporal = [
            np.clip(row_time / denom, 0.0, 1.0),
            np.clip(rel_age / denom, 0.0, 1.0),
            np.log1p(rel_age) / np.log1p(denom),
        ]
        time_values = np.stack(temporal[:remaining], axis=1).astype(np.float32)
        base = np.concatenate([base, time_values], axis=1)
    return base.astype(np.float32, copy=False)


def lagged_targets(target: np.ndarray, anchor: np.ndarray, lag_steps: int) -> np.ndarray:
    lagged = np.zeros((target.shape[0], target.shape[1], lag_steps), dtype=np.float32)
    for b in range(target.shape[0]):
        for i in range(target.shape[1]):
            for lag in range(lag_steps):
                j = i - lag - 1
                lagged[b, i, lag] = target[b, j] if j >= 0 and anchor[b, j] < anchor[b, i] else 0.0
    return lagged
