from __future__ import annotations

import argparse
import json
from dataclasses import replace
from pathlib import Path
from time import time

import numpy as np
import pandas as pd
import torch
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import mean_absolute_error, roc_auc_score

from .eval import logits_from_batch, multiclass_mrr, sigmoid_np
from .model import RelationalFoundationModel
from .relational_io import RelationalDataset, TASK_TYPES, temporal_train_end
from .targets import PAPER_TARGETS, compare_manifest_to_target
from .train import json_safe


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="In-context learning evaluation of a pre-trained KumoRFM model on RelBench tasks.")
    p.add_argument("--checkpoint", type=Path, required=True, help="Pre-trained model checkpoint.")
    p.add_argument("--dataset-dir", type=Path, required=True, help="Directory with exported RelBench task (metadata.json + parquet files).")
    p.add_argument("--output", type=Path, required=True, help="Output JSON path.")
    p.add_argument("--device", default="cuda", help="Evaluation device.")
    p.add_argument("--context-size", type=int, default=32, help="Max in-context examples.")
    p.add_argument("--batch-size", type=int, default=32, help="Number of queries to evaluate per batch.")
    p.add_argument("--rows-per-child", type=int, default=12)
    p.add_argument("--rows-per-aux", type=int, default=8)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--child-cols", type=int, default=8)
    p.add_argument("--aux-cols", type=int, default=6)
    p.add_argument("--lag-steps", type=int, default=4)
    p.add_argument("--max-eval-samples", type=int, default=2048, help="Maximum evaluation samples to process.")
    p.add_argument("--eval-min-target-classes", type=int, default=1, help="Minimum target classes required in the chronological eval prefix.")
    p.add_argument("--eval-diversity-scan-limit", type=int, default=0, help="Optional max rows to scan when extending eval prefix for target diversity; 0 disables extension.")
    p.add_argument("--ensemble-permutations", type=int, default=0, help="Average predictions over K random feature-column permutations (KumoRFM ensembling).")
    p.add_argument("--train-frac", type=float, default=0.7, help="Fraction of data for context (eval on remainder).")
    p.add_argument("--context-mode", choices=["recent", "entity", "entity_mixed"], default="recent")
    p.add_argument("--local-context-frac", type=float, default=0.5)
    p.add_argument("--root-baseline-max-train", type=int, default=0, help="Cap root-only baseline train rows; 0 uses all strict history.")
    p.add_argument("--indexed-table-lookup", action="store_true", help="Build per-table group indexes; useful for large eval windows.")
    p.add_argument("--skip-root-baseline", action="store_true", help="Skip root-only baseline fitting for fast model-only diagnostics.")
    p.add_argument(
        "--context-ablation",
        action="append",
        choices=["zero_targets", "shuffle_targets"],
        default=[],
        help="Negative-control ICL pass. zero_targets removes visible history labels; shuffle_targets permutes visible labels and zeros lagged labels.",
    )
    return p.parse_args()


@torch.no_grad()
def _ensemble_column_permutations(model, batch, n_perm: int, seed: int) -> torch.Tensor:
    """Average logits over random feature-column permutations (KumoRFM ensembling).

    The paper reduces variance and improves robustness by ensembling over
    randomized column/class orderings. Each table's feature columns are permuted
    consistently across all context+query samples, so within-task relational
    structure is preserved while the column ordering varies.
    """
    g = torch.Generator(device="cpu")
    g.manual_seed(seed)
    acc = logits_from_batch(model, batch).float()
    for _ in range(n_perm - 1):
        root = batch.root
        child = batch.child
        aux = batch.aux
        rp = torch.randperm(root.shape[-1], generator=g)
        cp = torch.randperm(child.shape[-1], generator=g)
        ap = torch.randperm(aux.shape[-1], generator=g)
        perm_batch = replace(
            batch,
            root=root[..., rp.to(root.device)],
            child=child[..., cp.to(child.device)],
            aux=aux[..., ap.to(aux.device)],
        )
        acc = acc + logits_from_batch(model, perm_batch).float()
    return acc / float(n_perm)


@torch.no_grad()
def evaluate_icl(
    model: RelationalFoundationModel,
    dataset: RelationalDataset,
    args: argparse.Namespace,
    device: torch.device,
) -> dict[str, float]:
    model.eval()
    task_spec = dataset.specs["task"]
    time_col = task_spec["time_column"]
    task_type = TASK_TYPES[task_spec["task_type"]]
    sorted_task = dataset.task.sort_values(time_col).reset_index(drop=True)
    train_end = temporal_train_end(sorted_task, time_col, train_frac=args.train_frac, context_size=args.context_size)
    if train_end <= 0 or train_end >= len(sorted_task):
        return {"error": "no valid strict temporal split"}
    train_rows = sorted_task.iloc[:train_end]
    eval_rows = sorted_task.iloc[train_end:]
    train_max_time = float(train_rows[time_col].max())
    eval_time = float(eval_rows[time_col].min())
    if train_max_time >= eval_time:
        return {"error": "temporal split is not strict", "train_max_time": train_max_time, "eval_min_time": eval_time}
    eval_start = int(len(train_rows))
    num_eval = choose_num_eval(eval_rows, task_spec["target_column"], task_type, args)
    if num_eval < 1:
        return {"error": "no evaluation samples after temporal split"}

    outputs = []
    targets = []
    y_classes = []
    num_classes_list = []
    ablation_outputs: dict[str, list[np.ndarray]] = {name: [] for name in getattr(args, "context_ablation", []) or []}
    num_batches = (num_eval + args.batch_size - 1) // args.batch_size

    for batch_idx in range(num_batches):
        start = eval_start + batch_idx * args.batch_size
        end = min(start + args.batch_size, eval_start + num_eval)
        if start >= end:
            break
        batch = dataset.to_batch(
            context_size=args.context_size,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            pred_start=start,
            batch_size=end - start,
            context_mode=getattr(args, "context_mode", "recent"),
            local_context_frac=getattr(args, "local_context_frac", 0.5),
            indexed_table_lookup=bool(getattr(args, "indexed_table_lookup", False)),
        ).to(device)
        n_perm = int(getattr(args, "ensemble_permutations", 0) or 0)
        if n_perm > 1:
            out = _ensemble_column_permutations(model, batch, n_perm, seed=1000 + batch_idx).cpu().numpy()
        else:
            out = logits_from_batch(model, batch).float().cpu().numpy()
        out = np.nan_to_num(out, nan=0.0, posinf=1e4, neginf=-1e4)
        outputs.append(out)
        for ablation_name in ablation_outputs:
            ablated = ablate_context_targets(batch, ablation_name, seed=batch_idx)
            ablated_out = logits_from_batch(model, ablated).float().cpu().numpy()
            ablated_out = np.nan_to_num(ablated_out, nan=0.0, posinf=1e4, neginf=-1e4)
            ablation_outputs[ablation_name].append(ablated_out)
        targets.append(batch.y.cpu().numpy())
        y_classes.append(batch.y_class.cpu().numpy())
        num_classes_list.append(batch.num_classes.cpu().numpy())

    if not outputs:
        return {"error": "no batches collected"}

    output = np.concatenate(outputs)
    y = np.concatenate(targets)
    y_class = np.concatenate(y_classes)
    n_classes = np.concatenate(num_classes_list)

    eval_times = eval_rows[time_col].iloc[:num_eval]
    unique_eval_times = eval_times.nunique()
    class_counts = dict(pd.Series(y).value_counts().to_dict()) if task_type <= 1 else {}
    metrics: dict[str, float] = {
        "num_eval_samples": float(len(y)),
        "requested_max_eval_samples": float(args.max_eval_samples),
        "eval_diversity_scan_limit": float(getattr(args, "eval_diversity_scan_limit", 0) or 0),
        "eval_min_target_classes": float(getattr(args, "eval_min_target_classes", 1) or 1),
        "num_context_rows": float(len(train_rows)),
        "train_max_time": train_max_time,
        "eval_min_time": eval_time,
        "eval_max_time": float(eval_rows[time_col].iloc[min(num_eval - 1, len(eval_rows) - 1)]),
        "eval_unique_timestamps": float(unique_eval_times),
        "train_unique_timestamps": float(train_rows[time_col].nunique()),
        "eval_class_distribution": str(class_counts),
    }

    if task_type == 0:
        if len(np.unique(y)) > 1:
            metrics["binary_auroc"] = float(roc_auc_score(y, sigmoid_np(output[:, 0])))
        else:
            metrics["binary_auroc"] = float("nan")
    elif task_type == 1:
        metrics["regression_mae"] = float(mean_absolute_error(y, output[:, 1]))
    elif task_type == 2:
        n_class_val = int(n_classes[0]) if len(n_classes) else 2
        if n_class_val > 2:
            metrics["multiclass_mrr"] = multiclass_mrr(output[:, 2:], y_class, n_classes)
        else:
            if len(np.unique(y)) > 1:
                metrics["binary_auroc"] = float(roc_auc_score(y, sigmoid_np(output[:, 0])))
            else:
                metrics["binary_auroc"] = float("nan")

    for ablation_name, ablated_chunks in ablation_outputs.items():
        if ablated_chunks:
            metrics.update(metric_delta_for_ablation(ablation_name, np.concatenate(ablated_chunks), output, y, y_class, n_classes, task_type))

    metrics["eval_batches"] = float(len(outputs))

    root_features = dataset.specs["tables"][0].get("feature_columns", [])
    if not bool(getattr(args, "skip_root_baseline", False)) and root_features and len(y) > 1:
        try:
            root_only = evaluate_root_only_baseline(dataset, eval_start, num_eval, args)
            metrics.update(root_only)
        except Exception as exc:
            metrics["root_only_error"] = repr(exc)
    elif bool(getattr(args, "skip_root_baseline", False)):
        metrics["root_only_skipped"] = True

    return metrics


def choose_num_eval(eval_rows: pd.DataFrame, target_col: str, task_type: int, args: argparse.Namespace) -> int:
    requested = min(int(args.max_eval_samples), len(eval_rows))
    if requested <= 0:
        return 0
    min_classes = int(getattr(args, "eval_min_target_classes", 1) or 1)
    scan_limit = int(getattr(args, "eval_diversity_scan_limit", 0) or 0)
    if task_type > 0 or min_classes <= 1 or scan_limit <= requested:
        return requested
    scan_limit = min(scan_limit, len(eval_rows))
    targets = eval_rows[target_col].iloc[:scan_limit]
    prefix_unique = targets.iloc[:requested].nunique(dropna=False)
    if prefix_unique >= min_classes:
        return requested
    for end in range(requested + 1, scan_limit + 1):
        if targets.iloc[:end].nunique(dropna=False) >= min_classes:
            return end
    return requested


def ablate_context_targets(batch, mode: str, seed: int) -> object:
    context_target = batch.context_target.clone()
    lag_target = batch.lag_target.clone()
    valid_context = (~batch.query_mask) & batch.table_mask.any(dim=-1)
    if mode == "zero_targets":
        context_target = context_target.masked_fill(valid_context, 0.0)
        lag_target = torch.zeros_like(lag_target)
    elif mode == "shuffle_targets":
        generator = torch.Generator(device=context_target.device)
        generator.manual_seed(17_003 + seed)
        for row_idx in range(context_target.shape[0]):
            positions = torch.nonzero(valid_context[row_idx], as_tuple=False).flatten()
            if positions.numel() > 1:
                perm = positions[torch.randperm(positions.numel(), generator=generator, device=context_target.device)]
                context_target[row_idx, positions] = context_target[row_idx, perm]
        # Lag features encode recent target history too; zero them so unshuffled
        # target sequences cannot leak through the negative control.
        lag_target = torch.zeros_like(lag_target)
    else:
        raise ValueError(f"Unknown context ablation: {mode}")
    return replace(batch, context_target=context_target, lag_target=lag_target)


def metric_delta_for_ablation(
    name: str,
    ablated_output: np.ndarray,
    output: np.ndarray,
    y: np.ndarray,
    y_class: np.ndarray,
    n_classes: np.ndarray,
    task_type: int,
) -> dict[str, float]:
    prefix = f"ablation_{name}"
    metrics: dict[str, float] = {}
    if task_type == 0:
        if len(np.unique(y)) > 1:
            ablated = float(roc_auc_score(y, sigmoid_np(ablated_output[:, 0])))
            base = float(roc_auc_score(y, sigmoid_np(output[:, 0])))
            metrics[f"{prefix}_binary_auroc"] = ablated
            metrics[f"{prefix}_binary_auroc_delta"] = base - ablated
        else:
            metrics[f"{prefix}_binary_auroc"] = float("nan")
            metrics[f"{prefix}_binary_auroc_delta"] = float("nan")
    elif task_type == 1:
        ablated_mae = float(mean_absolute_error(y, ablated_output[:, 1]))
        base_mae = float(mean_absolute_error(y, output[:, 1]))
        metrics[f"{prefix}_regression_mae"] = ablated_mae
        metrics[f"{prefix}_regression_mae_delta"] = ablated_mae - base_mae
    elif task_type == 2:
        n_class_val = int(n_classes[0]) if len(n_classes) else 2
        if n_class_val > 2:
            ablated_mrr = multiclass_mrr(ablated_output[:, 2:], y_class, n_classes)
            base_mrr = multiclass_mrr(output[:, 2:], y_class, n_classes)
            metrics[f"{prefix}_multiclass_mrr"] = ablated_mrr
            metrics[f"{prefix}_multiclass_mrr_delta"] = base_mrr - ablated_mrr
        elif len(np.unique(y)) > 1:
            ablated = float(roc_auc_score(y, sigmoid_np(ablated_output[:, 0])))
            base = float(roc_auc_score(y, sigmoid_np(output[:, 0])))
            metrics[f"{prefix}_binary_auroc"] = ablated
            metrics[f"{prefix}_binary_auroc_delta"] = base - ablated
        else:
            metrics[f"{prefix}_binary_auroc"] = float("nan")
            metrics[f"{prefix}_binary_auroc_delta"] = float("nan")
    return metrics


def evaluate_root_only_baseline(dataset: RelationalDataset, eval_start: int, num_eval: int, args: argparse.Namespace) -> dict[str, float]:
    task_spec = dataset.specs["task"]
    time_col = task_spec["time_column"]
    target_col = task_spec["target_column"]
    entity_key = task_spec["entity_key"]
    root_spec = dataset.specs["tables"][0]
    root_features = [f for f in root_spec.get("feature_columns", [])[:8] if f in dataset.root.columns]
    if not root_features:
        return {}
    sorted_task = dataset.task.sort_values(time_col).reset_index(drop=True)
    train = sorted_task.iloc[:eval_start]
    max_train = int(getattr(args, "root_baseline_max_train", 0) or 0)
    if max_train > 0 and len(train) > max_train:
        train = train.tail(max_train)
    eval_df = sorted_task.iloc[eval_start:eval_start + num_eval]
    X_train = root_feature_matrix(dataset, train, root_spec, root_features, entity_key, time_col)
    X_eval = root_feature_matrix(dataset, eval_df, root_spec, root_features, entity_key, time_col)
    y_train = train[target_col].to_numpy(dtype=np.float32)
    y_eval = eval_df[target_col].to_numpy(dtype=np.float32)
    if task_spec["task_type"] == "regression":
        from sklearn.ensemble import HistGradientBoostingRegressor

        reg = HistGradientBoostingRegressor(random_state=0)
        reg.fit(X_train, y_train)
        pred = reg.predict(X_eval)
        return {"root_only_mae": float(mean_absolute_error(y_eval, pred)), "root_only_train_rows": float(len(train))}
    if len(np.unique(y_train)) < 2 or len(np.unique(y_eval)) < 2:
        return {}
    clf = LogisticRegression(max_iter=500, class_weight="balanced")
    clf.fit(X_train, y_train.astype(int))
    proba = clf.predict_proba(X_eval)
    score = proba[:, 1] if proba.shape[1] > 1 else proba[:, 0]
    return {"root_only_auroc": float(roc_auc_score(y_eval, score)), "root_only_train_rows": float(len(train))}


def root_feature_matrix(
    dataset: RelationalDataset,
    rows: pd.DataFrame,
    root_spec: dict,
    root_features: list[str],
    entity_key: str,
    time_col: str,
) -> np.ndarray:
    values = []
    root_pk = root_spec["primary_key"]
    root_time = root_spec.get("time_column")
    root = dataset.root
    if not root_time or root_time not in root.columns:
        root_index = root.drop_duplicates(root_pk, keep="last").set_index(root_pk)
        matrix = root_index.reindex(rows[entity_key])[root_features]
        return matrix.apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=np.float32)
    if root_time and root_time in root.columns:
        root = root.sort_values(root_time)
    for _, row in rows.iterrows():
        candidates = root[root[root_pk] == row[entity_key]]
        if root_time and root_time in candidates.columns:
            candidates = candidates[candidates[root_time] < row[time_col]]
        if len(candidates):
            vals = pd.to_numeric(candidates[root_features].tail(1).iloc[0], errors="coerce").fillna(0.0).to_numpy(dtype=np.float32)
        else:
            vals = np.zeros(len(root_features), dtype=np.float32)
        values.append(vals)
    if not values:
        return np.zeros((0, len(root_features)), dtype=np.float32)
    return np.stack(values).astype(np.float32, copy=False)


def load_model_from_checkpoint(checkpoint_path: Path, device: torch.device) -> RelationalFoundationModel:
    checkpoint = torch.load(checkpoint_path, map_location=device, weights_only=False)
    manifest = checkpoint.get("manifest", {})
    train_args = manifest.get("args", {})
    state_dict = checkpoint["model"]
    readout_weight = None
    for key in ("readout.6.weight", "readout.3.weight", "readout.4.weight", "readout.5.weight"):
        if key in state_dict:
            readout_weight = state_dict[key]
            break
    if readout_weight is not None:
        out_dim = readout_weight.shape[0]
        max_classes = max(2, out_dim - 2)
    else:
        max_classes = 8
    model = RelationalFoundationModel(
        root_cols=int(train_args.get("root_cols", 8)),
        child_cols=int(train_args.get("child_cols", 8)),
        aux_cols=int(train_args.get("aux_cols", 6)),
        lag_steps=int(train_args.get("lag_steps", 4)),
        max_classes=max_classes,
        d_model=int(train_args.get("d_model", 256)),
        heads=int(train_args.get("heads", 8)),
        table_layers=int(train_args.get("table_layers") or train_args.get("layers", 3)),
        graph_layers=int(train_args.get("graph_layers") or train_args.get("layers", 3)),
        use_improved=not bool(train_args.get("no_improved")),
        typed_graph_attention=bool(train_args.get("typed_graph_attention", False)),
        typed_task_conditioning=bool(train_args.get("typed_task_conditioning", False)),
        row_bridge_attention=bool(train_args.get("row_bridge_attention", False)),
    ).to(device)
    model.load_state_dict(state_dict)
    return model


def run_icl_suite(
    checkpoint_path: Path,
    data_dir: Path,
    output_dir: Path,
    task_specs: list[dict],
    device: torch.device,
    **kwargs,
) -> dict:
    model = load_model_from_checkpoint(checkpoint_path, device)
    results = []
    for spec in task_specs:
        dataset_dir = data_dir / spec["dir_name"]
        if not (dataset_dir / "metadata.json").exists():
            results.append({"task": spec["name"], "ok": False, "error": "dataset not found"})
            continue
        dataset = RelationalDataset.load(dataset_dir)
        args_ns = argparse.Namespace(**{**kwargs, "dataset_dir": str(dataset_dir)})
        metrics = evaluate_icl(model, dataset, args_ns, device)
        results.append({"task": spec["name"], "ok": "error" not in metrics, **metrics})
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = {
        "checkpoint": str(checkpoint_path),
        "device": str(device),
        "tasks": results,
        "summary": icl_summary(results),
    }
    (output_dir / "icl_eval_summary.json").write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    return payload


def icl_summary(results: list[dict]) -> dict:
    ok_tasks = [t for t in results if t.get("ok", False)]
    aurocs = []
    maes = []
    mrrs = []
    for t in ok_tasks:
        if "binary_auroc" in t and np.isfinite(t["binary_auroc"]):
            aurocs.append(float(t["binary_auroc"]))
        if "regression_mae" in t and np.isfinite(t["regression_mae"]):
            maes.append(float(t["regression_mae"]))
        if "multiclass_mrr" in t and np.isfinite(t["multiclass_mrr"]):
            mrrs.append(float(t["multiclass_mrr"]))
    return {
        "num_tasks": len(results),
        "num_ok": len(ok_tasks),
        "avg_auroc": float(np.mean(aurocs)) if aurocs else None,
        "avg_auroc_points": float(np.mean(aurocs) * 100.0) if aurocs else None,
        "avg_mae": float(np.mean(maes)) if maes else None,
        "avg_mrr": float(np.mean(mrrs)) if mrrs else None,
    }


def main() -> None:
    args = parse_args()
    device = torch.device(args.device if torch.cuda.is_available() or args.device == "cpu" else "cpu")
    started = time()
    dataset = RelationalDataset.load(args.dataset_dir)
    model = load_model_from_checkpoint(args.checkpoint, device)
    metrics = evaluate_icl(model, dataset, args, device)
    payload = {
        "checkpoint": str(args.checkpoint),
        "dataset_dir": str(args.dataset_dir),
        "device": str(device),
        "elapsed_sec": time() - started,
        "metrics": metrics,
        "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    print(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
