from __future__ import annotations

import argparse
import importlib.util
import json
import os
from contextlib import nullcontext
from pathlib import Path
from time import time

import numpy as np
import pandas as pd
import torch
import torch.distributed as dist
import torch.nn.functional as F
from sklearn.metrics import mean_absolute_error, roc_auc_score
from torch.nn.parallel import DistributedDataParallel
from tqdm import trange

from .eval import logits_from_batch, multiclass_mrr, sigmoid_np
from .model import RelationalFoundationModel
from .relational_io import RelationalDataset, evaluate_flattened_baseline, temporal_train_end, write_demo_dataset
from .targets import compare_manifest_to_target, targets_payload
from .train import build_scheduler, json_safe


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="File-backed relational benchmark adapter smoke utilities.")
    sub = p.add_subparsers(dest="cmd", required=True)

    demo = sub.add_parser("write-demo", help="Write a small relational demo dataset.")
    demo.add_argument("--dataset-dir", type=Path, required=True)
    demo.add_argument("--rows", type=int, default=256)
    demo.add_argument("--seed", type=int, default=0)

    baseline = sub.add_parser("baseline", help="Evaluate a flattened baseline.")
    baseline.add_argument("--dataset-dir", type=Path, required=True)
    baseline.add_argument("--output", type=Path, default=None)

    validate = sub.add_parser("validate", help="Validate file-backed relational metadata and temporal split.")
    validate.add_argument("--dataset-dir", type=Path, required=True)
    validate.add_argument("--train-frac", type=float, default=0.7)
    validate.add_argument("--context-size", type=int, default=16)
    validate.add_argument("--output", type=Path, default=None)

    batch = sub.add_parser("batch-smoke", help="Materialize a RelationalBatch and run the current model forward.")
    batch.add_argument("--dataset-dir", type=Path, required=True)
    batch.add_argument("--context-size", type=int, default=16)
    batch.add_argument("--batch-size", type=int, default=8)
    batch.add_argument("--rows-per-child", type=int, default=8)
    batch.add_argument("--rows-per-aux", type=int, default=8)
    batch.add_argument("--root-cols", type=int, default=8)
    batch.add_argument("--child-cols", type=int, default=8)
    batch.add_argument("--aux-cols", type=int, default=6)
    batch.add_argument("--lag-steps", type=int, default=4)
    batch.add_argument("--d-model", type=int, default=32)
    batch.add_argument("--heads", type=int, default=4)
    batch.add_argument("--layers", type=int, default=1)
    batch.add_argument("--typed-task-conditioning", action=argparse.BooleanOptionalAction, default=False)
    batch.add_argument("--row-bridge-attention", action=argparse.BooleanOptionalAction, default=False)
    batch.add_argument("--output", type=Path, default=None)

    train = sub.add_parser("train-model", help="Fine-tune the relational model on a file-backed dataset.")
    train.add_argument("--dataset-dir", type=Path, required=True)
    train.add_argument("--output-dir", type=Path, required=True)
    train.add_argument("--steps", type=int, default=100)
    train.add_argument("--batch-size", type=int, default=16)
    train.add_argument("--context-size", type=int, default=16)
    train.add_argument("--rows-per-child", type=int, default=8)
    train.add_argument("--rows-per-aux", type=int, default=8)
    train.add_argument("--root-cols", type=int, default=8)
    train.add_argument("--child-cols", type=int, default=8)
    train.add_argument("--aux-cols", type=int, default=6)
    train.add_argument("--lag-steps", type=int, default=4)
    train.add_argument("--d-model", type=int, default=256)
    train.add_argument("--heads", type=int, default=8)
    train.add_argument("--layers", type=int, default=3)
    train.add_argument("--table-layers", type=int, default=None)
    train.add_argument("--graph-layers", type=int, default=None)
    train.add_argument("--no-improved", action="store_true")
    train.add_argument("--typed-task-conditioning", action=argparse.BooleanOptionalAction, default=False)
    train.add_argument("--row-bridge-attention", action=argparse.BooleanOptionalAction, default=False)
    train.add_argument("--lr", type=float, default=3e-4)
    train.add_argument("--weight-decay", type=float, default=1e-2)
    train.add_argument("--grad-accum-steps", type=int, default=1)
    train.add_argument("--warmup-steps", type=int, default=0)
    train.add_argument("--min-lr-ratio", type=float, default=0.1)
    train.add_argument("--train-frac", type=float, default=0.7)
    train.add_argument("--eval-batches", type=int, default=8)
    train.add_argument("--eval-min-rows", type=int, default=64, help="Minimum temporal evaluation rows to collect when rows are available.")
    train.add_argument("--eval-max-batches", type=int, default=128, help="Upper bound on temporal evaluation batches; allows small eval-batches to extend until metrics are computable.")
    train.add_argument("--balanced-loss", action=argparse.BooleanOptionalAction, default=True, help="Use train-split class-balanced losses for classification tasks.")
    train.add_argument("--balanced-batches", action=argparse.BooleanOptionalAction, default=False, help="Prefer classification train windows containing multiple labels.")
    train.add_argument("--amp", action="store_true")
    train.add_argument("--seed", type=int, default=0)
    train.add_argument("--save-checkpoint", action="store_true", default=True)
    train.add_argument("--checkpoint-every", type=int, default=0)
    train.add_argument("--resume", type=Path, default=None)
    train.add_argument("--log-every", type=int, default=1)

    targets = sub.add_parser("targets", help="Print registered paper target metrics.")
    targets.add_argument("--output", type=Path, default=None)

    compare = sub.add_parser("compare-target", help="Compare a manifest metric against a registered paper target.")
    compare.add_argument("--manifest", type=Path, required=True)
    compare.add_argument("--target-key", required=True)
    compare.add_argument("--metric-path", default=None)
    compare.add_argument("--output", type=Path, default=None)

    export = sub.add_parser("relbench-export", help="List or export RelBench tasks into the file-backed adapter format.")
    export.add_argument("--dataset", default=None, help="RelBench dataset name, for example rel-amazon.")
    export.add_argument("--task", default=None, help="RelBench task name, for example user-churn.")
    export.add_argument("--split", default="train", choices=("train", "val", "test"))
    export.add_argument("--output-dir", type=Path, default=None, help="Directory for metadata.json and parquet tables.")
    export.add_argument("--output", type=Path, default=None)
    export.add_argument("--list", action="store_true", help="Only list available datasets and tasks.")
    export.add_argument("--download", action="store_true", help="Use official RelBench downloads for database/task tables.")
    export.add_argument("--root-table", default=None)
    export.add_argument("--child-table", default=None)
    export.add_argument("--aux-table", default=None)
    export.add_argument("--max-rows", type=int, default=200_000, help="Maximum rows to keep per exported source table.")
    export.add_argument("--max-link-classes", type=int, default=512, help="Keep the most frequent destination classes when exporting link tasks.")
    export.add_argument("--max-link-positives-per-row", type=int, default=4, help="Maximum positive destination ids to explode per source timestamp.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if args.cmd == "write-demo":
        write_demo_dataset(args.dataset_dir, rows=args.rows, seed=args.seed)
        print(f"wrote {args.dataset_dir}")
        return

    if args.cmd == "targets":
        write_or_print({"targets": targets_payload()}, args.output)
        return

    if args.cmd == "compare-target":
        write_or_print(compare_manifest_to_target(args.manifest, args.target_key, args.metric_path), args.output)
        return

    if args.cmd == "relbench-export":
        payload = relbench_export(args)
        write_or_print(payload, args.output)
        return

    dataset = RelationalDataset.load(args.dataset_dir)
    if args.cmd == "baseline":
        metrics = evaluate_flattened_baseline(dataset)
        payload = {"dataset_dir": str(args.dataset_dir), "baseline": metrics}
        write_or_print(payload, args.output)
        return

    if args.cmd == "validate":
        payload = {"dataset_dir": str(args.dataset_dir), "validation": dataset.validate(train_frac=args.train_frac, context_size=args.context_size)}
        write_or_print(payload, args.output)
        return

    if args.cmd == "batch-smoke":
        batch = dataset.to_batch(
            context_size=args.context_size,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            batch_size=args.batch_size,
        )
        model = RelationalFoundationModel(
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            lag_steps=args.lag_steps,
            max_classes=int(batch.num_classes.max().item()),
            d_model=args.d_model,
            heads=args.heads,
            table_layers=args.layers,
            graph_layers=args.layers,
            typed_task_conditioning=args.typed_task_conditioning,
            row_bridge_attention=args.row_bridge_attention,
        )
        with torch.no_grad():
            output = model(
                batch.root,
                batch.child,
                batch.aux,
                batch.child_mask,
                batch.aux_mask,
                batch.table_mask,
                batch.context_target,
                batch.lag_target,
                batch.time,
                batch.query_mask,
            )
        payload = {
            "dataset_dir": str(args.dataset_dir),
            "batch": {
                "root": list(batch.root.shape),
                "child": list(batch.child.shape),
                "aux": list(batch.aux.shape),
                "task_type": batch.task_type.tolist(),
                "y": batch.y.tolist(),
            },
            "model_output_shape": list(output.shape),
        }
        write_or_print(payload, args.output)
        return

    if args.cmd == "train-model":
        train_file_model(args)


def train_file_model(args: argparse.Namespace) -> None:
    distributed, rank, world, local_rank = setup_distributed()
    device = torch.device(f"cuda:{local_rank}" if torch.cuda.is_available() else "cpu")
    torch.manual_seed(args.seed + rank)
    rng = np.random.default_rng(args.seed + rank * 10_000)
    started = time()

    dataset = RelationalDataset.load(args.dataset_dir)
    validation = dataset.validate(train_frac=args.train_frac, context_size=args.context_size)
    if not validation["ok"]:
        raise ValueError(f"Dataset validation failed: {validation['errors']}")
    task_spec = dataset.specs["task"]
    train_end = temporal_train_end(dataset.task, task_spec["time_column"], train_frac=args.train_frac, context_size=args.context_size)
    loss_weights = compute_file_loss_weights(dataset, train_end, device, balanced=args.balanced_loss)
    train_sampler = build_train_sampler(dataset, train_end, args.batch_size, args.context_size, balanced=args.balanced_batches)
    max_classes = int(task_spec.get("num_classes", 2))
    table_layers = args.table_layers if args.table_layers is not None else args.layers
    graph_layers = args.graph_layers if args.graph_layers is not None else args.layers
    model = RelationalFoundationModel(
        root_cols=args.root_cols,
        child_cols=args.child_cols,
        aux_cols=args.aux_cols,
        lag_steps=args.lag_steps,
        max_classes=max_classes,
        d_model=args.d_model,
        heads=args.heads,
        table_layers=table_layers,
        graph_layers=graph_layers,
        use_improved=not args.no_improved,
        typed_task_conditioning=args.typed_task_conditioning,
        row_bridge_attention=args.row_bridge_attention,
    ).to(device)
    if distributed:
        model = DistributedDataParallel(model, device_ids=[local_rank])
    opt = torch.optim.AdamW(model.parameters(), lr=args.lr, weight_decay=args.weight_decay)
    scheduler = build_scheduler(opt, args.steps, args.warmup_steps, args.min_lr_ratio)
    scaler = torch.amp.GradScaler("cuda", enabled=args.amp and device.type == "cuda")
    raw_model = model.module if distributed else model
    start_step = 0
    if args.resume is not None:
        checkpoint = torch.load(args.resume, map_location=device, weights_only=False)
        raw_model.load_state_dict(checkpoint["model"])
        if "optimizer" in checkpoint:
            opt.load_state_dict(checkpoint["optimizer"])
        if "scheduler" in checkpoint:
            scheduler.load_state_dict(checkpoint["scheduler"])
        if "scaler" in checkpoint:
            scaler.load_state_dict(checkpoint["scaler"])
        start_step = int(checkpoint.get("global_step", 0))
        if "rng_state" in checkpoint:
            rng.bit_generator.state = checkpoint["rng_state"]
        if "torch_rng_state" in checkpoint:
            torch.set_rng_state(checkpoint["torch_rng_state"].cpu())
        if torch.cuda.is_available() and checkpoint.get("cuda_rng_state_all") is not None:
            torch.cuda.set_rng_state_all([state.cpu() for state in checkpoint["cuda_rng_state_all"]])

    model.train()
    train_started = time()
    iterator = trange(start_step, args.steps, disable=rank != 0)
    last_loss = float("nan")
    skipped_nonfinite_steps = 0
    completed_steps = 0
    args.grad_accum_steps = max(1, args.grad_accum_steps)
    for step in iterator:
        opt.zero_grad(set_to_none=True)
        accum_loss = 0.0
        finite_step = True
        for accum_idx in range(args.grad_accum_steps):
            pred_start = sample_train_start(rng, train_sampler)
            batch = dataset.to_batch(
                context_size=args.context_size,
                rows_per_child=args.rows_per_child,
                rows_per_aux=args.rows_per_aux,
                lag_steps=args.lag_steps,
                root_cols=args.root_cols,
                child_cols=args.child_cols,
                aux_cols=args.aux_cols,
                pred_start=pred_start,
                batch_size=args.batch_size,
            ).to(device)
            sync_context = nullcontext()
            if distributed and accum_idx < args.grad_accum_steps - 1:
                sync_context = model.no_sync()
            with sync_context:
                ctx = torch.amp.autocast("cuda", enabled=args.amp) if device.type == "cuda" else nullcontext()
                with ctx:
                    output = logits_from_batch(model, batch)
                    loss = file_task_loss(output, batch, loss_weights)
                finite_loss = torch.isfinite(loss).float()
                if distributed:
                    dist.all_reduce(finite_loss, op=dist.ReduceOp.MIN)
                if finite_loss.item() < 1.0:
                    finite_step = False
                    last_loss = float(loss.detach().cpu()) if rank == 0 else last_loss
                    break
                scaled_loss = loss / args.grad_accum_steps
                scaler.scale(scaled_loss).backward()
                accum_loss += float(loss.detach().cpu())
        if not finite_step:
            skipped_nonfinite_steps += 1
            opt.zero_grad(set_to_none=True)
            if rank == 0:
                iterator.set_postfix(loss="nonfinite-skip", world=world)
            continue
        scaler.unscale_(opt)
        finite_grad = finite_gradients(raw_model)
        if distributed:
            dist.all_reduce(finite_grad, op=dist.ReduceOp.MIN)
        if finite_grad.item() < 1.0:
            skipped_nonfinite_steps += 1
            opt.zero_grad(set_to_none=True)
            scaler.update()
            if rank == 0:
                iterator.set_postfix(loss="nonfinite-grad-skip", world=world)
            continue
        torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
        old_scale = scaler.get_scale()
        scaler.step(opt)
        scaler.update()
        if not scaler.is_enabled() or scaler.get_scale() >= old_scale:
            scheduler.step()
        completed_steps += 1
        last_loss = accum_loss / args.grad_accum_steps
        if rank == 0 and (step + 1) % max(1, args.log_every) == 0:
            iterator.set_postfix(loss=f"{last_loss:.4f}", lr=f"{scheduler.get_last_lr()[0]:.2e}", world=world)
        if rank == 0 and args.checkpoint_every > 0 and (step + 1) % args.checkpoint_every == 0:
            args.output_dir.mkdir(parents=True, exist_ok=True)
            manifest = {
                "dataset_dir": str(args.dataset_dir),
                "world_size": world,
                "global_step": step + 1,
                "last_train_loss": last_loss,
                "skipped_nonfinite_steps": skipped_nonfinite_steps,
                "learning_rate": scheduler.get_last_lr()[0],
            }
            save_file_checkpoint(args.output_dir / f"checkpoint_step_{step + 1}.pt", raw_model, opt, scheduler, scaler, manifest, step + 1, rng)

    train_elapsed = max(1e-9, time() - train_started)
    metrics = evaluate_file_model(raw_model, dataset, args, device, train_end)
    if rank == 0:
        baseline = evaluate_flattened_baseline(dataset, train_frac=args.train_frac)
        args.output_dir.mkdir(parents=True, exist_ok=True)
        effective_batch_size = args.batch_size * world * args.grad_accum_steps
        manifest = {
            "dataset_dir": str(args.dataset_dir),
            "world_size": world,
            "device": str(device),
            "elapsed_sec": time() - started,
            "train_elapsed_sec": train_elapsed,
            "effective_batch_size": effective_batch_size,
            "optimizer_steps": completed_steps,
            "global_step": args.steps,
            "resumed_from_step": start_step,
            "learning_rate": scheduler.get_last_lr()[0],
            "train_tasks_per_sec": completed_steps * effective_batch_size / train_elapsed,
            "train_context_examples_per_sec": completed_steps * effective_batch_size * (args.context_size + 1) / train_elapsed,
            "last_train_loss": last_loss,
            "skipped_nonfinite_steps": skipped_nonfinite_steps,
            "model_metrics": metrics,
            "flattened_baseline": baseline,
            "loss_weights": json_safe(loss_weight_manifest(loss_weights)),
            "train_sampler": json_safe(train_sampler_manifest(train_sampler)),
            "dataset_validation": validation,
            "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
        }
        (args.output_dir / "manifest.json").write_text(json.dumps(json_safe(manifest), allow_nan=False, indent=2, sort_keys=True) + "\n")
        if args.save_checkpoint:
            save_file_checkpoint(args.output_dir / "checkpoint.pt", raw_model, opt, scheduler, scaler, manifest, args.steps, rng)
    if distributed:
        dist.destroy_process_group()


def compute_file_loss_weights(dataset: RelationalDataset, train_end: int, device: torch.device, balanced: bool) -> dict[str, torch.Tensor | None]:
    if not balanced:
        return {"binary_pos_weight": None, "multiclass_weight": None}
    task_spec = dataset.specs["task"]
    target_col = task_spec["target_column"]
    task_type = task_spec["task_type"]
    train_task = dataset.task.sort_values(task_spec["time_column"]).reset_index(drop=True).iloc[:train_end]
    if task_type == "binary":
        y = pd.to_numeric(train_task[target_col], errors="coerce").dropna().astype(float)
        positives = float((y > 0.5).sum())
        negatives = float((y <= 0.5).sum())
        if positives > 0.0 and negatives > 0.0:
            return {"binary_pos_weight": torch.tensor([negatives / positives], dtype=torch.float32, device=device), "multiclass_weight": None}
    if task_type == "multiclass":
        num_classes = int(task_spec.get("num_classes", 2))
        y = pd.to_numeric(train_task[target_col], errors="coerce").dropna().astype(int)
        counts = np.bincount(y[(0 <= y) & (y < num_classes)], minlength=num_classes).astype(np.float32)
        if np.count_nonzero(counts) > 1:
            weights = counts.sum() / np.maximum(counts, 1.0)
            weights = weights / max(float(weights.mean()), 1e-6)
            return {"binary_pos_weight": None, "multiclass_weight": torch.tensor(weights, dtype=torch.float32, device=device)}
    return {"binary_pos_weight": None, "multiclass_weight": None}


def loss_weight_manifest(loss_weights: dict[str, torch.Tensor | None]) -> dict[str, object]:
    out: dict[str, object] = {}
    binary = loss_weights.get("binary_pos_weight")
    out["binary_pos_weight"] = None if binary is None else float(binary.detach().cpu().reshape(-1)[0])
    multiclass = loss_weights.get("multiclass_weight")
    out["multiclass_weight"] = None if multiclass is None else [float(v) for v in multiclass.detach().cpu().tolist()]
    return out


def build_train_sampler(dataset: RelationalDataset, train_end: int, batch_size: int, context_size: int, balanced: bool) -> dict[str, object]:
    low = int(context_size)
    high = max(low, int(train_end) - int(batch_size))
    fallback = np.arange(low, high + 1, dtype=np.int64)
    task_spec = dataset.specs["task"]
    if not balanced or task_spec.get("task_type") not in {"binary", "multiclass"} or len(fallback) == 0:
        return {"starts": fallback, "fallback": fallback, "balanced": False, "num_balanced_starts": 0, "num_fallback_starts": int(len(fallback))}
    target_col = task_spec["target_column"]
    time_col = task_spec["time_column"]
    task = dataset.task.sort_values(time_col).reset_index(drop=True)
    y = pd.to_numeric(task[target_col], errors="coerce").fillna(-1).astype(int)
    starts = [int(start) for start in fallback if y.iloc[start : start + batch_size].nunique() > 1]
    balanced_starts = np.array(starts, dtype=np.int64)
    return {
        "starts": balanced_starts if len(balanced_starts) else fallback,
        "fallback": fallback,
        "balanced": bool(len(balanced_starts)),
        "num_balanced_starts": int(len(balanced_starts)),
        "num_fallback_starts": int(len(fallback)),
    }


def sample_train_start(rng: np.random.Generator, sampler: dict[str, object]) -> int:
    starts = sampler["starts"]
    if not isinstance(starts, np.ndarray) or len(starts) == 0:
        fallback = sampler["fallback"]
        starts = fallback if isinstance(fallback, np.ndarray) and len(fallback) else np.array([0], dtype=np.int64)
    return int(starts[int(rng.integers(0, len(starts)))])


def train_sampler_manifest(sampler: dict[str, object]) -> dict[str, object]:
    return {
        "balanced": bool(sampler.get("balanced")),
        "num_balanced_starts": int(sampler.get("num_balanced_starts", 0)),
        "num_fallback_starts": int(sampler.get("num_fallback_starts", 0)),
    }


def file_task_loss(output: torch.Tensor, batch, loss_weights: dict[str, torch.Tensor | None]) -> torch.Tensor:
    output = output.float()
    losses: list[torch.Tensor] = []
    binary = batch.task_type == 0
    if binary.any():
        losses.append(F.binary_cross_entropy_with_logits(output[binary, 0], batch.y[binary].float(), pos_weight=loss_weights.get("binary_pos_weight")))
    regression = batch.task_type == 1
    if regression.any():
        losses.append(F.mse_loss(output[regression, 1], batch.y[regression]))
    multiclass = batch.task_type == 2
    if multiclass.any():
        class_logits = output[multiclass, 2:]
        class_targets = batch.y_class[multiclass]
        class_counts = batch.num_classes[multiclass]
        class_ids = torch.arange(class_logits.shape[1], device=class_logits.device).unsqueeze(0)
        inactive = class_ids >= class_counts.unsqueeze(1)
        class_logits = class_logits.masked_fill(inactive, -1e4)
        weight = loss_weights.get("multiclass_weight")
        if weight is not None:
            weight = weight[: class_logits.shape[1]]
        losses.append(F.cross_entropy(class_logits, class_targets, weight=weight))
    if not losses:
        return output.sum() * 0.0
    return torch.stack(losses).mean()


def finite_gradients(model: torch.nn.Module) -> torch.Tensor:
    device = next(model.parameters()).device
    finite = torch.ones((), device=device)
    for param in model.parameters():
        if param.grad is not None and not torch.isfinite(param.grad).all():
            finite = torch.zeros((), device=device)
            break
    return finite


@torch.no_grad()
def evaluate_file_model(
    model: RelationalFoundationModel,
    dataset: RelationalDataset,
    args: argparse.Namespace,
    device: torch.device,
    train_end: int,
) -> dict[str, float]:
    was_training = model.training
    model.eval()
    outputs = []
    targets = []
    classes = []
    task_types = []
    num_classes = []
    max_batches = max(args.eval_batches, args.eval_max_batches)
    for offset in range(max_batches):
        pred_start = train_end + offset * args.batch_size
        if pred_start >= len(dataset.task):
            break
        batch = dataset.to_batch(
            context_size=args.context_size,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            pred_start=pred_start,
            batch_size=args.batch_size,
        ).to(device)
        out = logits_from_batch(model, batch).float().cpu().numpy()
        outputs.append(np.nan_to_num(out, nan=0.0, posinf=1e4, neginf=-1e4))
        targets.append(batch.y.cpu().numpy())
        classes.append(batch.y_class.cpu().numpy())
        task_types.append(batch.task_type.cpu().numpy())
        num_classes.append(batch.num_classes.cpu().numpy())
        if offset + 1 >= args.eval_batches and eval_collection_sufficient(targets, task_types, args.eval_min_rows):
            break
    if was_training:
        model.train()
    if not outputs:
        return {}
    output = np.concatenate(outputs)
    y = np.concatenate(targets)
    y_class = np.concatenate(classes)
    task_type = np.concatenate(task_types)
    n_classes = np.concatenate(num_classes)
    metrics: dict[str, float] = {}
    if task_type[0] == 0 and len(np.unique(y)) > 1:
        metrics["binary_auroc"] = float(roc_auc_score(y, sigmoid_np(output[:, 0])))
    elif task_type[0] == 1:
        metrics["regression_mae"] = float(mean_absolute_error(y, output[:, 1]))
    elif task_type[0] == 2:
        metrics["multiclass_mrr"] = multiclass_mrr(output[:, 2:], y_class, n_classes)
    metrics["eval_rows"] = float(len(y))
    metrics["eval_batches_used"] = float(len(outputs))
    metrics["eval_unique_targets"] = float(len(np.unique(y if task_type[0] != 2 else y_class)))
    metrics.update(evaluate_link_ranking_model(model, dataset, args, device, train_end))
    return metrics


def eval_collection_sufficient(targets: list[np.ndarray], task_types: list[np.ndarray], eval_min_rows: int) -> bool:
    y = np.concatenate(targets)
    task_type = int(np.concatenate(task_types)[0])
    if len(y) < eval_min_rows:
        return False
    if task_type == 0:
        return len(np.unique(y)) > 1
    return True


@torch.no_grad()
def evaluate_link_ranking_model(
    model: RelationalFoundationModel,
    dataset: RelationalDataset,
    args: argparse.Namespace,
    device: torch.device,
    train_end: int,
) -> dict[str, float]:
    source = dataset.specs.get("source", {})
    if source.get("link_export") != "positive_destination_multiclass":
        return {}
    link_target_path = source.get("link_target_path")
    if not link_target_path:
        return {}
    path = Path(args.dataset_dir) / link_target_path
    if not path.exists():
        return {}

    task_spec = dataset.specs["task"]
    entity_key = task_spec["entity_key"]
    time_col = task_spec["time_column"]
    target_col = task_spec["target_column"]
    num_classes = int(task_spec["num_classes"])
    k = min(10, num_classes)
    sorted_task = dataset.task.sort_values(time_col).reset_index(drop=True)
    if train_end < len(sorted_task):
        eval_start_time = float(sorted_task.iloc[train_end][time_col])
    else:
        eval_start_time = float(sorted_task[time_col].max())

    link_targets = pd.read_parquet(path).sort_values(time_col).reset_index(drop=True)
    eval_rows = link_targets[link_targets[time_col] >= eval_start_time].head(args.eval_batches * args.batch_size)
    if len(eval_rows) == 0:
        eval_rows = link_targets.tail(args.eval_batches * args.batch_size)
    if len(eval_rows) == 0:
        return {}

    reciprocal_ranks = []
    hits = []
    recalls = []
    for _, row in eval_rows.iterrows():
        true_classes = normalize_class_list(row["target_classes"])
        if not true_classes:
            continue
        anchor = float(row[time_col])
        history = sorted_task[sorted_task[time_col] < anchor].tail(args.context_size)
        query = pd.DataFrame([{entity_key: row[entity_key], time_col: anchor, target_col: true_classes[0]}])
        eval_task = pd.concat([history, query], ignore_index=True)
        eval_dataset = RelationalDataset(dataset.root, dataset.child, dataset.aux, eval_task, dataset.specs)
        batch = eval_dataset.to_batch(
            context_size=args.context_size,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            pred_start=max(0, len(eval_task) - 1),
            batch_size=1,
        ).to(device)
        logits = logits_from_batch(model, batch)[0, 2 : 2 + num_classes].detach().cpu().numpy()
        order = np.argsort(-logits)
        true_set = set(true_classes)
        top_order = order[:k]
        first_rank = next((rank + 1 for rank, cls in enumerate(top_order) if int(cls) in true_set), None)
        topk = {int(cls) for cls in top_order}
        hits.append(float(bool(true_set & topk)))
        recalls.append(len(true_set & topk) / len(true_set))
        reciprocal_ranks.append(0.0 if first_rank is None else 1.0 / first_rank)
    if not reciprocal_ranks:
        return {}
    return {
        f"link_mrr_at_{k}": float(np.mean(reciprocal_ranks)),
        f"link_hit_at_{k}": float(np.mean(hits)),
        f"link_recall_at_{k}": float(np.mean(recalls)),
        "link_eval_rows": float(len(reciprocal_ranks)),
    }


def setup_distributed() -> tuple[bool, int, int, int]:
    if "RANK" not in os.environ:
        return False, 0, 1, 0
    dist.init_process_group(backend="nccl")
    rank = dist.get_rank()
    world = dist.get_world_size()
    local_rank = int(os.environ.get("LOCAL_RANK", 0))
    torch.cuda.set_device(local_rank)
    return True, rank, world, local_rank


def save_file_checkpoint(
    path: Path,
    raw_model,
    opt: torch.optim.Optimizer,
    scheduler,
    scaler: torch.amp.GradScaler,
    manifest: dict[str, object],
    global_step: int,
    rng: np.random.Generator,
) -> None:
    torch.save(
        {
            "model": raw_model.state_dict(),
            "optimizer": opt.state_dict(),
            "scheduler": scheduler.state_dict(),
            "scaler": scaler.state_dict(),
            "manifest": manifest,
            "global_step": global_step,
            "rng_state": rng.bit_generator.state,
            "torch_rng_state": torch.get_rng_state(),
            "cuda_rng_state_all": torch.cuda.get_rng_state_all() if torch.cuda.is_available() else None,
        },
        path,
    )


def relbench_export(args: argparse.Namespace) -> dict:
    missing = [module for module in ("relbench", "torch_frame") if importlib.util.find_spec(module) is None]
    if missing:
        return {
            "ok": False,
            "missing_dependencies": missing,
            "next_step": "Install official benchmark dependencies, then implement dataset/task-specific export into metadata.json + parquet tables.",
        }
    from relbench.datasets import get_dataset, get_dataset_names
    from relbench.tasks import get_task, get_task_names

    dataset_names = get_dataset_names()
    task_names = {name: get_task_names(name) for name in dataset_names}
    if args.list or args.dataset is None or args.task is None:
        return {
            "ok": True,
            "missing_dependencies": [],
            "datasets": dataset_names,
            "tasks": task_names,
            "next_step": "Run relbench-export with --dataset, --task, --output-dir, and --download when official cached tables are required.",
        }
    if args.output_dir is None:
        return {"ok": False, "error": "--output-dir is required when exporting a dataset/task."}
    if args.dataset not in task_names:
        return {"ok": False, "error": f"Unknown RelBench dataset {args.dataset!r}.", "datasets": dataset_names}
    if args.task not in task_names[args.dataset]:
        return {"ok": False, "error": f"Unknown task {args.task!r} for {args.dataset!r}.", "tasks": task_names[args.dataset]}

    dataset = get_dataset(args.dataset, download=args.download)
    task = get_task(args.dataset, args.task, download=args.download)
    task_kind = str(getattr(task, "task_type", ""))
    entity_table = getattr(task, "entity_table", None)
    entity_col = getattr(task, "entity_col", None)
    time_col = getattr(task, "time_col", None)
    target_col = getattr(task, "target_col", None)
    adapter_task_type = relbench_task_type_to_adapter(getattr(task, "task_type", None))
    db = dataset.get_db(upto_test_timestamp=True)
    task_table = task.get_table(args.split, mask_input_cols=False)

    if adapter_task_type is not None and all([entity_table, entity_col, time_col, target_col]):
        export_payload = export_relbench_entity_task(
            db=db,
            task_df=task_table.df,
            output_dir=args.output_dir,
            root_table=args.root_table or entity_table,
            entity_col=entity_col,
            time_col=time_col,
            target_col=target_col,
            adapter_task_type=adapter_task_type,
            child_table=args.child_table,
            aux_table=args.aux_table,
            max_rows=args.max_rows,
        )
    elif getattr(getattr(task, "task_type", None), "value", None) == "link_prediction":
        export_payload = export_relbench_link_task(
            db=db,
            task_df=task_table.df,
            output_dir=args.output_dir,
            src_table=args.root_table or getattr(task, "src_entity_table"),
            src_col=getattr(task, "src_entity_col"),
            dst_table=getattr(task, "dst_entity_table"),
            dst_col=getattr(task, "dst_entity_col"),
            time_col=time_col,
            child_table=args.child_table,
            aux_table=args.aux_table,
            max_rows=args.max_rows,
            max_link_classes=args.max_link_classes,
            max_link_positives_per_row=args.max_link_positives_per_row,
        )
    else:
        return {"ok": False, "dataset": args.dataset, "task": args.task, "task_type": task_kind, "error": "Unsupported task shape for current adapter."}
    export_payload.update(
        {
            "ok": True,
            "dataset": args.dataset,
            "task": args.task,
            "split": args.split,
            "download": bool(args.download),
            "relbench_task_type": task_kind,
        }
    )
    return export_payload


def relbench_task_type_to_adapter(task_type) -> str | None:
    value = getattr(task_type, "value", str(task_type))
    if value == "binary_classification":
        return "binary"
    if value == "regression":
        return "regression"
    if value == "multiclass_classification":
        return "multiclass"
    return None


def export_relbench_entity_task(
    db,
    task_df: pd.DataFrame,
    output_dir: Path,
    root_table: str,
    entity_col: str,
    time_col: str,
    target_col: str,
    adapter_task_type: str,
    child_table: str | None,
    aux_table: str | None,
    max_rows: int,
) -> dict:
    table_dict = db.table_dict
    if root_table not in table_dict:
        raise ValueError(f"Root table {root_table!r} is not present in RelBench database.")

    root_src = table_dict[root_table]
    # Select child/aux by ACTUAL CONNECTIVITY to the prediction entities, not raw
    # table size. A table can be huge yet rarely link to the task entities (e.g.
    # rel-stack `votes` has 1.3M rows but connects to almost none of the badge-task
    # users, while `badges` connects to 463k). Using distinct-entity coverage is a
    # principled relational-graph criterion (prefer tables that link to what we
    # predict), not benchmark-specific tuning, and it fixes wrong-child selection
    # uniformly across star schemas.
    task_entities = set(task_df[entity_col].dropna().tolist())
    coverage_children = choose_children_by_coverage(table_dict, root_table, task_entities, top_k=2)
    child_name = child_table or (coverage_children[0] if coverage_children else choose_child_table(table_dict, root_table))
    aux_name = aux_table or (coverage_children[1] if len(coverage_children) > 1 else (choose_child_table(table_dict, child_name) if child_name else None))
    root_df, root_pk, root_time, root_features = prepare_table(root_src.df, root_src.pkey_col, root_src.time_col, root_src.fkey_col_to_pkey_table, max_rows)
    task_out = prepare_task_table(task_df, entity_col, time_col, target_col, adapter_task_type)

    root_ids = set(root_df[root_pk].dropna().tolist())
    task_out = task_out[task_out[entity_col].isin(root_ids)].sort_values(time_col).reset_index(drop=True)
    if len(task_out) == 0:
        raise ValueError("No task rows remain after filtering to exported root entities.")

    if child_name:
        child_src = table_dict[child_name]
        child_fk = find_fk_to_table(child_src.fkey_col_to_pkey_table, root_table)
        child_df, child_pk, child_time, child_features = prepare_table(child_src.df, child_src.pkey_col, child_src.time_col, child_src.fkey_col_to_pkey_table, max_rows)
        if child_fk is None:
            child_fk = "__root_id"
            child_df[child_fk] = np.nan
        child_df = child_df[child_df[child_fk].isin(root_ids)].reset_index(drop=True)
    else:
        child_name = "relbench_empty_child"
        child_pk = "event_id"
        child_fk = root_pk
        child_time = time_col
        child_features = ["child_bias"]
        child_df = pd.DataFrame({child_pk: pd.Series(dtype="int64"), child_fk: pd.Series(dtype=root_df[root_pk].dtype), child_time: pd.Series(dtype="float64"), "child_bias": pd.Series(dtype="float32")})

    # aux parent: "child" for chain schemas (aux FK->child); "root" for star schemas
    # where aux is a second sibling child (aux FK->root). The star fallback lets the
    # model see a second of N sibling tables instead of an empty aux — the dominant
    # structure in rel-stack / rel-amazon, where the 3-table-chain assumption fails.
    aux_parent = "child"
    aux_fk_to_child = find_fk_to_table(table_dict[aux_name].fkey_col_to_pkey_table, child_name) if (aux_name and aux_name in table_dict) else None
    # Prefer the highest-connectivity second sibling (coverage), falling back to size.
    star_aux = next((c for c in coverage_children if c != child_name), None) or choose_child_table(table_dict, root_table, exclude=child_name)
    if aux_name and len(child_df) and aux_fk_to_child:
        aux_src = table_dict[aux_name]
        aux_fk = aux_fk_to_child
        aux_df, aux_pk, aux_time, aux_features = prepare_table(aux_src.df, aux_src.pkey_col, aux_src.time_col, aux_src.fkey_col_to_pkey_table, max_rows)
        aux_df = aux_df[aux_df[aux_fk].isin(set(child_df[child_pk].dropna().tolist()))].reset_index(drop=True)
    elif star_aux and len(child_df):
        aux_parent = "root"
        aux_name = star_aux
        aux_src = table_dict[aux_name]
        aux_fk = find_fk_to_table(aux_src.fkey_col_to_pkey_table, root_table)
        aux_df, aux_pk, aux_time, aux_features = prepare_table(aux_src.df, aux_src.pkey_col, aux_src.time_col, aux_src.fkey_col_to_pkey_table, max_rows)
        aux_df = aux_df[aux_df[aux_fk].isin(root_ids)].reset_index(drop=True)
    else:
        aux_name = "relbench_empty_aux"
        aux_pk = "aux_id"
        aux_fk = child_pk
        aux_time = child_time
        aux_features = ["aux_bias"]
        aux_df = pd.DataFrame({aux_pk: pd.Series(dtype="int64"), aux_fk: pd.Series(dtype=child_df[child_pk].dtype), aux_time: pd.Series(dtype="float64"), "aux_bias": pd.Series(dtype="float32")})

    output_dir.mkdir(parents=True, exist_ok=True)
    root_df.to_parquet(output_dir / "root.parquet", index=False)
    child_df.to_parquet(output_dir / "child.parquet", index=False)
    aux_df.to_parquet(output_dir / "aux.parquet", index=False)
    task_out.to_parquet(output_dir / "task.parquet", index=False)
    metadata = {
        "root_table": "root",
        "child_table": "child",
        "aux_table": "aux",
        "tables": [
            {"name": "root", "path": "root.parquet", "primary_key": root_pk, "time_column": root_time, "feature_columns": root_features},
            {"name": "child", "path": "child.parquet", "primary_key": child_pk, "foreign_key": child_fk, "parent_table": "root", "time_column": child_time, "feature_columns": child_features},
            {"name": "aux", "path": "aux.parquet", "primary_key": aux_pk, "foreign_key": aux_fk, "parent_table": aux_parent, "time_column": aux_time, "feature_columns": aux_features},
        ],
        "task": {
            "path": "task.parquet",
            "entity_table": "root",
            "entity_key": entity_col,
            "time_column": time_col,
            "target_column": target_col,
            "task_type": adapter_task_type,
            "num_classes": int(task_out[target_col].nunique()) if adapter_task_type == "multiclass" else 2,
        },
        "source": {"root_table": root_table, "child_table": child_name, "aux_table": aux_name},
    }
    (output_dir / "metadata.json").write_text(json.dumps(json_safe(metadata), allow_nan=False, indent=2, sort_keys=True) + "\n")
    return {
        "output_dir": str(output_dir),
        "rows": {"root": len(root_df), "child": len(child_df), "aux": len(aux_df), "task": len(task_out)},
        "feature_counts": {"root": len(root_features), "child": len(child_features), "aux": len(aux_features)},
        "selected_tables": {"root": root_table, "child": child_name, "aux": aux_name},
    }


def export_relbench_link_task(
    db,
    task_df: pd.DataFrame,
    output_dir: Path,
    src_table: str,
    src_col: str,
    dst_table: str,
    dst_col: str,
    time_col: str,
    child_table: str | None,
    aux_table: str | None,
    max_rows: int,
    max_link_classes: int,
    max_link_positives_per_row: int,
) -> dict:
    table_dict = db.table_dict
    if src_table not in table_dict:
        raise ValueError(f"Source table {src_table!r} is not present in RelBench database.")
    if dst_table not in table_dict:
        raise ValueError(f"Destination table {dst_table!r} is not present in RelBench database.")

    src_df, src_pk, src_time, src_features = prepare_table(
        table_dict[src_table].df,
        table_dict[src_table].pkey_col,
        table_dict[src_table].time_col,
        table_dict[src_table].fkey_col_to_pkey_table,
        max_rows,
    )
    dst_values = task_df[dst_col].explode().dropna()
    class_values = dst_values.value_counts().head(max_link_classes).index.tolist()
    class_to_idx = {value: idx for idx, value in enumerate(class_values)}
    task_out = prepare_link_task_table(
        task_df,
        src_col=src_col,
        time_col=time_col,
        dst_col=dst_col,
        class_to_idx=class_to_idx,
        max_positives_per_row=max_link_positives_per_row,
    )
    src_ids = set(src_df[src_pk].dropna().tolist())
    task_out = task_out[task_out[src_col].isin(src_ids)].sort_values(time_col).reset_index(drop=True)
    if len(task_out) == 0:
        raise ValueError("No link task rows remain after filtering source entities and destination classes.")

    child_name = child_table or choose_child_table(table_dict, src_table)
    aux_name = aux_table or (choose_child_table(table_dict, child_name) if child_name else None)
    if child_name:
        child_src = table_dict[child_name]
        child_fk = find_fk_to_table(child_src.fkey_col_to_pkey_table, src_table)
        child_df, child_pk, child_time, child_features = prepare_table(child_src.df, child_src.pkey_col, child_src.time_col, child_src.fkey_col_to_pkey_table, max_rows)
        if child_fk is None:
            child_fk = "__root_id"
            child_df[child_fk] = np.nan
        child_df = child_df[child_df[child_fk].isin(src_ids)].reset_index(drop=True)
    else:
        child_name = "relbench_empty_child"
        child_pk = "event_id"
        child_fk = src_pk
        child_time = time_col
        child_features = ["child_bias"]
        child_df = pd.DataFrame({child_pk: pd.Series(dtype="int64"), child_fk: pd.Series(dtype=src_df[src_pk].dtype), child_time: pd.Series(dtype="float64"), "child_bias": pd.Series(dtype="float32")})

    if aux_name and len(child_df):
        aux_src = table_dict[aux_name]
        aux_fk = find_fk_to_table(aux_src.fkey_col_to_pkey_table, child_name)
        aux_df, aux_pk, aux_time, aux_features = prepare_table(aux_src.df, aux_src.pkey_col, aux_src.time_col, aux_src.fkey_col_to_pkey_table, max_rows)
        if aux_fk is None:
            aux_fk = "__child_id"
            aux_df[aux_fk] = np.nan
        aux_df = aux_df[aux_df[aux_fk].isin(set(child_df[child_pk].dropna().tolist()))].reset_index(drop=True)
    else:
        aux_name = "relbench_empty_aux"
        aux_pk = "aux_id"
        aux_fk = child_pk
        aux_time = child_time
        aux_features = ["aux_bias"]
        aux_df = pd.DataFrame({aux_pk: pd.Series(dtype="int64"), aux_fk: pd.Series(dtype=child_df[child_pk].dtype), aux_time: pd.Series(dtype="float64"), "aux_bias": pd.Series(dtype="float32")})

    output_dir.mkdir(parents=True, exist_ok=True)
    src_df.to_parquet(output_dir / "root.parquet", index=False)
    child_df.to_parquet(output_dir / "child.parquet", index=False)
    aux_df.to_parquet(output_dir / "aux.parquet", index=False)
    task_out.to_parquet(output_dir / "task.parquet", index=False)
    link_targets = prepare_link_target_table(
        task_df,
        src_col=src_col,
        time_col=time_col,
        dst_col=dst_col,
        class_to_idx=class_to_idx,
    )
    link_targets = link_targets[link_targets[src_col].isin(src_ids)].sort_values(time_col).reset_index(drop=True)
    link_targets.to_parquet(output_dir / "link_targets.parquet", index=False)
    class_map = [{"class_index": int(idx), "destination_id": scalar_json(value)} for value, idx in class_to_idx.items()]
    (output_dir / "class_map.json").write_text(json.dumps(class_map, allow_nan=False, indent=2, sort_keys=True) + "\n")
    metadata = {
        "root_table": "root",
        "child_table": "child",
        "aux_table": "aux",
        "tables": [
            {"name": "root", "path": "root.parquet", "primary_key": src_pk, "time_column": src_time, "feature_columns": src_features},
            {"name": "child", "path": "child.parquet", "primary_key": child_pk, "foreign_key": child_fk, "parent_table": "root", "time_column": child_time, "feature_columns": child_features},
            {"name": "aux", "path": "aux.parquet", "primary_key": aux_pk, "foreign_key": aux_fk, "parent_table": "child", "time_column": aux_time, "feature_columns": aux_features},
        ],
        "task": {
            "path": "task.parquet",
            "entity_table": "root",
            "entity_key": src_col,
            "time_column": time_col,
            "target_column": "target_class",
            "task_type": "multiclass",
            "num_classes": len(class_to_idx),
        },
        "source": {
            "root_table": src_table,
            "child_table": child_name,
            "aux_table": aux_name,
            "dst_table": dst_table,
            "dst_col": dst_col,
            "link_export": "positive_destination_multiclass",
            "link_target_path": "link_targets.parquet",
        },
    }
    (output_dir / "metadata.json").write_text(json.dumps(json_safe(metadata), allow_nan=False, indent=2, sort_keys=True) + "\n")
    return {
        "output_dir": str(output_dir),
        "rows": {"root": len(src_df), "child": len(child_df), "aux": len(aux_df), "task": len(task_out), "link_targets": len(link_targets)},
        "feature_counts": {"root": len(src_features), "child": len(child_features), "aux": len(aux_features)},
        "selected_tables": {"root": src_table, "child": child_name, "aux": aux_name, "destination": dst_table},
        "link_export": {"mode": "positive_destination_multiclass", "num_classes": len(class_to_idx), "max_positives_per_source_timestamp": max_link_positives_per_row},
    }


def choose_children_by_coverage(table_dict: dict, root_table: str, entity_ids: set, top_k: int = 2) -> list[str]:
    """Rank FK->root tables by how many prediction entities they actually connect to.

    Returns table names ordered by distinct-entity coverage (desc). This is a
    schema-graph criterion (use tables that link to the prediction entities),
    independent of any specific task target.
    """
    if not entity_ids:
        return []
    scored = []
    for name, table in table_dict.items():
        fk = find_fk_to_table(table.fkey_col_to_pkey_table, root_table)
        if fk is None or fk not in table.df.columns:
            continue
        vals = table.df[fk].dropna()
        if len(vals) == 0:
            continue
        # Sample for speed on huge tables; coverage is stable under sampling.
        sample = vals if len(vals) <= 500_000 else vals.sample(500_000, random_state=0)
        covered = len(set(sample.unique()) & entity_ids)
        if covered > 0:
            scored.append((covered, len(table.df), name))
    scored.sort(key=lambda x: (-x[0], -x[1]))
    return [name for _, _, name in scored[:top_k]]


def choose_child_table(table_dict: dict, parent_table: str | None, exclude: str | None = None) -> str | None:
    if parent_table is None:
        return None
    candidates = []
    for name, table in table_dict.items():
        if name == exclude:
            continue
        if find_fk_to_table(table.fkey_col_to_pkey_table, parent_table):
            candidates.append((table.time_col is None, -len(table.df), name))
    return sorted(candidates)[0][2] if candidates else None


def find_fk_to_table(fkeys: dict[str, str], parent_table: str) -> str | None:
    for col, table_name in fkeys.items():
        if table_name == parent_table:
            return col
    return None


def prepare_task_table(df: pd.DataFrame, entity_col: str, time_col: str, target_col: str, task_type: str) -> pd.DataFrame:
    out = df[[entity_col, time_col, target_col]].copy()
    out[time_col] = numeric_time(out[time_col])
    if task_type in {"binary", "multiclass"}:
        out[target_col] = pd.factorize(out[target_col], sort=True)[0].astype("int64")
    else:
        out[target_col] = pd.to_numeric(out[target_col], errors="coerce").fillna(0.0).astype("float32")
    return out.dropna(subset=[entity_col, time_col]).reset_index(drop=True)


def prepare_link_task_table(
    df: pd.DataFrame,
    src_col: str,
    time_col: str,
    dst_col: str,
    class_to_idx: dict,
    max_positives_per_row: int,
) -> pd.DataFrame:
    rows = []
    for _, row in df[[src_col, time_col, dst_col]].iterrows():
        src = row[src_col]
        timestamp = row[time_col]
        dst_values = row[dst_col]
        if not isinstance(dst_values, (list, tuple, set, np.ndarray, pd.Series)):
            dst_values = [dst_values]
        kept = 0
        for dst in dst_values:
            if dst not in class_to_idx:
                continue
            rows.append({src_col: src, time_col: timestamp, "target_class": class_to_idx[dst]})
            kept += 1
            if kept >= max_positives_per_row:
                break
    out = pd.DataFrame(rows, columns=[src_col, time_col, "target_class"])
    if len(out) == 0:
        return out
    out[time_col] = numeric_time(out[time_col])
    out["target_class"] = out["target_class"].astype("int64")
    return out.dropna(subset=[src_col, time_col]).reset_index(drop=True)


def prepare_link_target_table(
    df: pd.DataFrame,
    src_col: str,
    time_col: str,
    dst_col: str,
    class_to_idx: dict,
) -> pd.DataFrame:
    rows = []
    for _, row in df[[src_col, time_col, dst_col]].iterrows():
        classes = []
        dst_values = row[dst_col]
        if not isinstance(dst_values, (list, tuple, set, np.ndarray, pd.Series)):
            dst_values = [dst_values]
        for dst in dst_values:
            if dst in class_to_idx:
                classes.append(int(class_to_idx[dst]))
        if classes:
            rows.append({src_col: row[src_col], time_col: row[time_col], "target_classes": sorted(set(classes))})
    out = pd.DataFrame(rows, columns=[src_col, time_col, "target_classes"])
    if len(out) == 0:
        return out
    out[time_col] = numeric_time(out[time_col])
    return out.dropna(subset=[src_col, time_col]).reset_index(drop=True)


def normalize_class_list(value) -> list[int]:
    if isinstance(value, np.ndarray):
        return [int(v) for v in value.tolist()]
    if isinstance(value, (list, tuple, set, pd.Series)):
        return [int(v) for v in value]
    if pd.isna(value):
        return []
    return [int(value)]


def scalar_json(value):
    if isinstance(value, np.generic):
        return value.item()
    if pd.isna(value):
        return None
    return value


def prepare_table(df: pd.DataFrame, pkey_col: str | None, time_col: str | None, fkeys: dict[str, str], max_rows: int) -> tuple[pd.DataFrame, str, str | None, list[str]]:
    out = df.copy()
    if time_col and time_col in out.columns:
        out[time_col] = numeric_time(out[time_col])
        out = out.sort_values(time_col)
    if max_rows > 0 and len(out) > max_rows:
        out = out.tail(max_rows)
    if pkey_col is None or pkey_col not in out.columns:
        pkey_col = "__row_id"
        out[pkey_col] = np.arange(len(out), dtype=np.int64)
    keep = {pkey_col, *fkeys.keys()}
    if time_col:
        keep.add(time_col)
    feature_cols = []
    for col in list(out.columns):
        if col in keep:
            continue
        feature_cols.append(col)
        out[col] = encode_feature(out[col])
    if not feature_cols:
        feature_cols = ["__bias"]
        out["__bias"] = np.float32(1.0)
    return out[[pkey_col, *fkeys.keys(), *([time_col] if time_col else []), *feature_cols]].reset_index(drop=True), pkey_col, time_col, feature_cols


def numeric_time(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return (series.astype("int64") / 1e9).astype("float64")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0).astype("float64")
    converted = pd.to_datetime(series, errors="coerce")
    if converted.notna().mean() > 0.8:
        return (converted.astype("int64") / 1e9).astype("float64")
    return pd.to_numeric(series, errors="coerce").fillna(0.0).astype("float64")


def encode_feature(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return scale_feature(numeric_time(series))
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_bool_dtype(series):
        return scale_feature(pd.to_numeric(series, errors="coerce"))
    codes = pd.factorize(series.astype("string").fillna("<NA>"), sort=True)[0]
    return scale_feature(pd.Series(codes, index=series.index))


def scale_feature(series: pd.Series) -> pd.Series:
    values = pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0).astype("float64")
    std = float(values.std())
    if std > 1e-12:
        values = (values - float(values.mean())) / std
    else:
        values = values * 0.0
    return values.clip(-10.0, 10.0).astype("float32")


def write_or_print(payload: dict, output: Path | None) -> None:
    text = json.dumps(payload, indent=2, sort_keys=True) + "\n"
    if output is None:
        print(text, end="")
    else:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text)


if __name__ == "__main__":
    main()
