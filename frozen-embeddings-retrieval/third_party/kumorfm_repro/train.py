from __future__ import annotations

import argparse
import json
import math
import os
from contextlib import nullcontext
from pathlib import Path
from time import time

import numpy as np
import torch
import torch.distributed as dist
import torch.nn.functional as F
from torch import nn
from torch.nn.parallel import DistributedDataParallel
from tqdm import trange

from .data import RelationalTaskGenerator
from .real_data import RealTaskSampler
from .eval import (
    GeneratorConfig,
    evaluate_latent_structure,
    evaluate_mechanisms,
    evaluate_metrics,
    evaluate_robustness,
    logits_from_batch,
)
from .model import RelationalFoundationModel


def json_safe(value: object) -> object:
    if isinstance(value, float):
        return value if np.isfinite(value) else None
    if isinstance(value, dict):
        return {str(k): json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_safe(v) for v in value]
    return value


def finite_metric_summary(metrics: object) -> dict[str, object]:
    if not isinstance(metrics, dict):
        return {"mean": None, "num_finite": 0, "num_total": 0}
    values = [float(v) for v in metrics.values() if isinstance(v, (int, float))]
    finite = [v for v in values if np.isfinite(v)]
    return {
        "mean": float(np.mean(finite)) if finite else None,
        "num_finite": len(finite),
        "num_total": len(values),
    }


def eval_metric_summary(eval_metrics: dict[str, object]) -> dict[str, object]:
    summary: dict[str, object] = {}
    mechanisms = eval_metrics.get("mechanisms")
    if isinstance(mechanisms, dict):
        summary["mechanisms"] = {
            "mean": mechanisms.get("mean") if isinstance(mechanisms.get("mean"), (int, float)) else None,
            "num_mechanisms": len([k for k in mechanisms if str(k).startswith("mechanism_")]),
            "num_finite_mechanisms": len(
                [
                    v
                    for k, v in mechanisms.items()
                    if str(k).startswith("mechanism_") and isinstance(v, (int, float)) and np.isfinite(float(v))
                ]
            ),
        }
    robustness = eval_metrics.get("robustness")
    if isinstance(robustness, dict):
        summary["robustness"] = {
            "mean": robustness.get("mean") if isinstance(robustness.get("mean"), (int, float)) else None,
            "num_conditions": len([k for k in robustness if k != "mean"]),
            "num_finite_conditions": len(
                [
                    v
                    for k, v in robustness.items()
                    if k != "mean" and isinstance(v, (int, float)) and np.isfinite(float(v))
                ]
            ),
        }
    latent = eval_metrics.get("latent_structure")
    if isinstance(latent, dict):
        summary["latent_structure"] = finite_metric_summary(latent)
    return summary


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Train a KumoRFM-style relational foundation model on SCM tasks.")
    p.add_argument("--steps", type=int, default=1000)
    p.add_argument("--batch-size", type=int, default=16)
    p.add_argument("--context-size", type=int, default=64)
    p.add_argument("--rows-per-child", type=int, default=12)
    p.add_argument("--rows-per-aux", type=int, default=8)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--child-cols", type=int, default=8)
    p.add_argument("--aux-cols", type=int, default=6)
    p.add_argument("--lag-steps", type=int, default=4)
    p.add_argument("--d-model", type=int, default=384)
    p.add_argument("--heads", type=int, default=12)
    p.add_argument("--layers", type=int, default=4)
    p.add_argument("--table-layers", type=int, default=None, help="Override table encoder depth separately from graph layers.")
    p.add_argument("--graph-layers", type=int, default=None, help="Override graph encoder depth separately from table layers.")
    p.add_argument("--no-improved", action="store_true", help="Use baseline model instead of improved architecture.")
    p.add_argument(
        "--typed-graph-attention",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="Use pairwise typed table-to-table attention in the improved graph encoder.",
    )
    p.add_argument(
        "--typed-task-conditioning",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Add role embeddings for context target, visibility, lag target, and time conditioning columns.",
    )
    p.add_argument(
        "--row-bridge-attention",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="Add row-level child-to-aux cross-table attention before table-level pooling.",
    )
    p.add_argument("--lr", type=float, default=3e-4)
    p.add_argument("--weight-decay", type=float, default=1e-2)
    p.add_argument("--grad-accum-steps", type=int, default=1)
    p.add_argument("--warmup-steps", type=int, default=0)
    p.add_argument("--min-lr-ratio", type=float, default=0.1)
    p.add_argument("--missing-prob", type=float, default=0.05)
    p.add_argument("--link-dropout", type=float, default=0.0)
    p.add_argument("--schema-dropout", type=float, default=0.0)
    p.add_argument("--task-type", choices=["binary", "regression", "multiclass", "mixed"], default="binary")
    p.add_argument("--max-classes", type=int, default=8)
    p.add_argument("--binary-quantile-min", type=float, default=0.5, help="Lower score quantile for synthetic binary thresholds.")
    p.add_argument("--binary-quantile-max", type=float, default=0.5, help="Upper score quantile for synthetic binary thresholds.")
    p.add_argument("--temporal-row-features", action=argparse.BooleanOptionalAction, default=True, help="Inject row timestamp/age features into child and aux tables.")
    p.add_argument("--entity-history-frac", type=float, default=0.0, help="Fraction of context examples generated as same-entity history for the query.")
    p.add_argument("--entity-drift-scale", type=float, default=0.3, help="Temporal drift scale for synthetic same-entity latent trajectories.")
    p.add_argument("--root-shortcut-dropout", type=float, default=0.0, help="Probability of replacing synthetic root features with noise after target generation.")
    p.add_argument("--feature-heterogeneity", type=float, default=0.0, help="Probability of transforming synthetic columns into binary, ordinal, skewed, sparse, or count-like marginals.")
    p.add_argument("--history-label-frac", type=float, default=0.0, help="Fraction of synthetic tasks with same-entity autoregressive label dynamics.")
    p.add_argument("--relational-bridge-frac", type=float, default=0.0, help="Fraction of synthetic tasks with join-like child-to-aux bridge label dynamics.")
    p.add_argument("--feature-diversity", type=float, default=0.0, help="Fraction of features generated from non-normal distributions natively.")
    p.add_argument("--feature-correlation", type=float, default=0.0, help="Strength of random feature-level correlations within each table.")
    p.add_argument("--extended-mechanisms", action="store_true", help="Use 36 mechanisms instead of 24 for greater simulation diversity.")
    p.add_argument("--column-permutation", action="store_true", help="Per-task feature-column permutation for column-order-invariant relational reasoning.")
    p.add_argument("--real-data-dirs", type=str, nargs="*", default=[], help="Exported RelBench dataset dirs for real-database co-training (leave-one-dataset-out: never include the eval target).")
    p.add_argument("--real-data-frac", type=float, default=0.0, help="Probability a training batch is drawn from real datasets instead of synthetic.")
    p.add_argument("--real-train-frac", type=float, default=0.7, help="Train-region fraction (unique timestamps) anchors are sampled from in real co-training.")
    p.add_argument("--context-reconstruction-loss-weight", type=float, default=0.0, help="Auxiliary loss weight for predicting randomly masked visible context targets.")
    p.add_argument("--mechanism-aux-loss-weight", type=float, default=0.0, help="Pre-training-only auxiliary loss weight for synthetic mechanism prediction from query embeddings.")
    p.add_argument("--schema-aux-loss-weight", type=float, default=0.0, help="Pre-training-only auxiliary loss weight for table-availability prediction from query embeddings.")
    p.add_argument("--val-batches", type=int, default=32)
    p.add_argument("--eval-batches", type=int, default=16)
    p.add_argument("--eval-mechanisms", action="store_true")
    p.add_argument("--eval-robustness", action="store_true")
    p.add_argument("--eval-latent", action="store_true")
    p.add_argument("--num-workers", type=int, default=0, help="Reserved for future DataLoader-backed generators.")
    p.add_argument("--amp", action="store_true")
    p.add_argument("--amp-init-scale", type=float, default=1024.0, help="Initial GradScaler scale for AMP training.")
    p.add_argument("--amp-growth-interval", type=int, default=2000, help="GradScaler growth interval for AMP training.")
    p.add_argument("--ema-decay", type=float, default=0.0, help="EMA decay rate for model parameters (0=disabled, 0.999=slow, 0.99=fast). Smooths checkpoint instability.")
    p.add_argument("--seed", type=int, default=0)
    p.add_argument("--output-dir", type=Path, default=Path("runs/smoke"))
    p.add_argument("--save-checkpoint", action="store_true")
    p.add_argument("--checkpoint-every", type=int, default=0)
    p.add_argument("--validate-every", type=int, default=0, help="Run held-out synthetic validation every N training steps.")
    p.add_argument("--save-best-checkpoint", action="store_true", help="Save checkpoint_best.pt when periodic validation improves.")
    p.add_argument("--resume", type=Path, default=None)
    p.add_argument("--log-every", type=int, default=1)
    return p.parse_args()


def task_loss_with_targets(
    output: torch.Tensor,
    batch,
    y: torch.Tensor,
    y_class: torch.Tensor,
) -> torch.Tensor:
    output = output.float()
    losses: list[torch.Tensor] = []
    binary = batch.task_type == 0
    if binary.any():
        losses.append(F.binary_cross_entropy_with_logits(output[binary, 0], y[binary].float()))
    regression = batch.task_type == 1
    if regression.any():
        losses.append(F.mse_loss(output[regression, 1], y[regression]))
    multiclass = batch.task_type == 2
    if multiclass.any():
        class_logits = output[multiclass, 2:]
        class_targets = y_class[multiclass]
        class_counts = batch.num_classes[multiclass]
        class_ids = torch.arange(class_logits.shape[1], device=class_logits.device).unsqueeze(0)
        inactive = class_ids >= class_counts.unsqueeze(1)
        class_logits = class_logits.masked_fill(inactive, -1e4)
        losses.append(F.cross_entropy(class_logits, class_targets))
    if not losses:
        return output.sum() * 0.0
    return torch.stack(losses).mean()


def task_loss(output: torch.Tensor, batch) -> torch.Tensor:
    return task_loss_with_targets(output, batch, batch.y, batch.y_class)


def context_reconstruction_loss(
    model: nn.Module,
    batch,
    *,
    weight: float,
) -> tuple[torch.Tensor, dict[str, float]]:
    """Mask one visible context target per task and predict it from the rest."""

    if weight <= 0:
        return batch.root.sum() * 0.0, {}
    num_context = batch.context_target.shape[1] - 1
    if num_context <= 0:
        return batch.root.sum() * 0.0, {}

    sample_idx = torch.randint(0, num_context, (batch.root.shape[0],), device=batch.root.device)
    rows = torch.arange(batch.root.shape[0], device=batch.root.device)
    positions = torch.arange(batch.context_target.shape[1], device=batch.root.device).unsqueeze(0)
    future_or_query = positions >= sample_idx.unsqueeze(1)
    rec_context_target = batch.context_target.clone()
    rec_context_target = rec_context_target.masked_fill(future_or_query, 0.0)
    rec_query_mask = future_or_query

    rec_logits = model(
        batch.root,
        batch.child,
        batch.aux,
        batch.child_mask,
        batch.aux_mask,
        batch.table_mask,
        rec_context_target,
        batch.lag_target,
        batch.time,
        rec_query_mask,
    )
    rec_y = batch.target_all[rows, sample_idx]
    rec_y_class = batch.y_class_all[rows, sample_idx]
    rec_loss = task_loss_with_targets(rec_logits, batch, rec_y, rec_y_class)
    return rec_loss * weight, {"context_reconstruction_loss": float(rec_loss.detach().cpu())}


class StructuralAuxiliaryHeads(nn.Module):
    """Pre-training-only heads for relational structure supervision."""

    def __init__(self, d_model: int, num_mechanisms: int = 24, dropout: float = 0.1) -> None:
        super().__init__()
        self.shared = nn.Sequential(
            nn.LayerNorm(d_model),
            nn.Linear(d_model, d_model),
            nn.GELU(),
            nn.Dropout(dropout),
        )
        self.mechanism = nn.Linear(d_model, num_mechanisms)
        self.schema = nn.Linear(d_model, 3)

    def forward(self, embedding: torch.Tensor) -> tuple[torch.Tensor, torch.Tensor]:
        h = self.shared(embedding.float())
        return self.mechanism(h), self.schema(h)


def structural_auxiliary_loss(
    heads: nn.Module,
    embedding: torch.Tensor,
    batch,
    *,
    mechanism_weight: float,
    schema_weight: float,
) -> tuple[torch.Tensor, dict[str, float]]:
    mech_logits, schema_logits = heads(embedding)
    losses: list[torch.Tensor] = []
    metrics: dict[str, float] = {}
    if mechanism_weight > 0:
        mech_loss = F.cross_entropy(mech_logits, batch.mechanism)
        losses.append(mech_loss * mechanism_weight)
        metrics["mechanism_aux_loss"] = float(mech_loss.detach().cpu())
    if schema_weight > 0:
        schema_target = batch.table_mask[:, -1, :].float()
        schema_loss = F.binary_cross_entropy_with_logits(schema_logits, schema_target)
        losses.append(schema_loss * schema_weight)
        metrics["schema_aux_loss"] = float(schema_loss.detach().cpu())
    if not losses:
        return embedding.sum() * 0.0, metrics
    return torch.stack(losses).sum(), metrics


def setup_distributed() -> tuple[bool, int, int, int]:
    if "RANK" not in os.environ:
        return False, 0, 1, 0
    dist.init_process_group(backend="nccl")
    rank = dist.get_rank()
    world = dist.get_world_size()
    local_rank = int(os.environ.get("LOCAL_RANK", 0))
    torch.cuda.set_device(local_rank)
    return True, rank, world, local_rank


def build_scheduler(opt: torch.optim.Optimizer, total_steps: int, warmup_steps: int, min_lr_ratio: float):
    total_steps = max(1, total_steps)
    warmup_steps = max(0, min(warmup_steps, total_steps))
    min_lr_ratio = max(0.0, min(1.0, min_lr_ratio))

    def lr_lambda(step: int) -> float:
        if warmup_steps > 0 and step < warmup_steps:
            return max(1e-8, float(step + 1) / float(warmup_steps))
        progress_denom = max(1, total_steps - warmup_steps)
        progress = min(1.0, max(0.0, float(step - warmup_steps) / float(progress_denom)))
        cosine = 0.5 * (1.0 + math.cos(math.pi * progress))
        return min_lr_ratio + (1.0 - min_lr_ratio) * cosine

    return torch.optim.lr_scheduler.LambdaLR(opt, lr_lambda)


def finite_gradients(*models: nn.Module | None) -> torch.Tensor:
    first = next(model for model in models if model is not None)
    device = next(first.parameters()).device
    finite = torch.ones((), device=device)
    for model in models:
        if model is None:
            continue
        for param in model.parameters():
            if param.grad is not None and not torch.isfinite(param.grad).all():
                finite = torch.zeros((), device=device)
                break
        if finite.item() < 1.0:
            break
    return finite


class EMAModel:
    """Lightweight EMA wrapper that can temporarily swap model weights."""

    def __init__(self, ema_state: dict[str, torch.Tensor], model: nn.Module):
        self.ema_state = ema_state
        self.model = model
        self._backup: dict[str, torch.Tensor] = {}

    def __enter__(self):
        state = self.model.state_dict(keep_vars=True)
        self._backup = {k: v.detach().clone() for k, v in state.items()}
        self.model.load_state_dict(self.ema_state, strict=False)
        return self

    def __exit__(self, *args):
        self.model.load_state_dict(self._backup, strict=False)
        self._backup.clear()


def save_checkpoint(
    path: Path,
    raw_model: nn.Module,
    opt: torch.optim.Optimizer,
    scheduler,
    scaler: torch.amp.GradScaler,
    manifest: dict[str, object],
    global_step: int,
    train_gen: RelationalTaskGenerator,
    raw_aux_heads: nn.Module | None = None,
    ema_state: dict[str, torch.Tensor] | None = None,
) -> None:
    payload = {
        "model": raw_model.state_dict(),
        "optimizer": opt.state_dict(),
        "scheduler": scheduler.state_dict(),
        "scaler": scaler.state_dict(),
        "manifest": manifest,
        "global_step": global_step,
        "train_rng_state": train_gen.rng.bit_generator.state,
        "torch_rng_state": torch.get_rng_state(),
        "cuda_rng_state_all": torch.cuda.get_rng_state_all() if torch.cuda.is_available() else None,
    }
    if raw_aux_heads is not None:
        payload["auxiliary_heads"] = raw_aux_heads.state_dict()
    if ema_state is not None:
        payload["ema_model"] = {k: v.cpu().clone() for k, v in ema_state.items()}
    torch.save(payload, path)


def distributed_finite_mean(score: float, device: torch.device, distributed: bool) -> float:
    if not distributed:
        return score
    finite = float(np.isfinite(score))
    payload = torch.tensor([score if finite else 0.0, finite], device=device, dtype=torch.float64)
    dist.all_reduce(payload, op=dist.ReduceOp.SUM)
    count = float(payload[1].item())
    return float(payload[0].item() / count) if count > 0 else float("nan")


def main() -> None:
    args = parse_args()
    started = time()
    distributed, rank, world, local_rank = setup_distributed()
    device = torch.device(f"cuda:{local_rank}" if torch.cuda.is_available() else "cpu")
    torch.manual_seed(args.seed + rank)
    np.random.seed(args.seed + rank)

    train_gen = RelationalTaskGenerator(
        context_size=args.context_size,
        root_cols=args.root_cols,
        child_cols=args.child_cols,
        aux_cols=args.aux_cols,
        rows_per_child=args.rows_per_child,
        rows_per_aux=args.rows_per_aux,
        lag_steps=args.lag_steps,
        missing_prob=args.missing_prob,
        link_dropout=args.link_dropout,
        schema_dropout=args.schema_dropout,
        task_type=args.task_type,
        max_classes=args.max_classes,
        binary_quantile_min=args.binary_quantile_min,
        binary_quantile_max=args.binary_quantile_max,
        temporal_row_features=args.temporal_row_features,
        entity_history_frac=args.entity_history_frac,
        entity_drift_scale=args.entity_drift_scale,
        root_shortcut_dropout=args.root_shortcut_dropout,
        feature_heterogeneity=args.feature_heterogeneity,
        history_label_frac=args.history_label_frac,
        relational_bridge_frac=args.relational_bridge_frac,
        feature_diversity=args.feature_diversity,
        feature_correlation=args.feature_correlation,
        extended_mechanisms=args.extended_mechanisms,
        column_permutation=args.column_permutation,
        seed=args.seed + 10_000 * rank,
    )
    real_sampler = None
    real_dirs_flat = [p for entry in args.real_data_dirs for p in str(entry).split()]
    if real_dirs_flat and args.real_data_frac > 0:
        real_sampler = RealTaskSampler(
            [Path(d) for d in real_dirs_flat],
            context_size=args.context_size,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            train_frac=args.real_train_frac,
            seed=args.seed + 7919 * (rank + 1),
        )
        if rank == 0:
            print(f"real co-training datasets: {len(real_sampler)} | frac={args.real_data_frac}", flush=True)

    val_gen = RelationalTaskGenerator(
        context_size=args.context_size,
        root_cols=args.root_cols,
        child_cols=args.child_cols,
        aux_cols=args.aux_cols,
        rows_per_child=args.rows_per_child,
        rows_per_aux=args.rows_per_aux,
        lag_steps=args.lag_steps,
        missing_prob=args.missing_prob,
        link_dropout=args.link_dropout,
        schema_dropout=args.schema_dropout,
        task_type=args.task_type,
        max_classes=args.max_classes,
        binary_quantile_min=args.binary_quantile_min,
        binary_quantile_max=args.binary_quantile_max,
        temporal_row_features=args.temporal_row_features,
        entity_history_frac=args.entity_history_frac,
        entity_drift_scale=args.entity_drift_scale,
        root_shortcut_dropout=args.root_shortcut_dropout,
        feature_heterogeneity=args.feature_heterogeneity,
        history_label_frac=args.history_label_frac,
        relational_bridge_frac=args.relational_bridge_frac,
        feature_diversity=args.feature_diversity,
        feature_correlation=args.feature_correlation,
        extended_mechanisms=args.extended_mechanisms,
        column_permutation=args.column_permutation,
        seed=args.seed + 999,
    )

    table_layers = args.table_layers if args.table_layers is not None else args.layers
    graph_layers = args.graph_layers if args.graph_layers is not None else args.layers
    if args.typed_graph_attention is None:
        args.typed_graph_attention = True
        if args.resume is not None:
            resume_meta = torch.load(args.resume, map_location="cpu", weights_only=False)
            resume_args = resume_meta.get("manifest", {}).get("args", {})
            if "typed_graph_attention" in resume_args:
                args.typed_graph_attention = bool(resume_args["typed_graph_attention"])
            else:
                args.typed_graph_attention = any(
                    key.startswith("graph_encoder.typed_graph_blocks.") for key in resume_meta.get("model", {})
                )
    model = RelationalFoundationModel(
        root_cols=args.root_cols,
        child_cols=args.child_cols,
        aux_cols=args.aux_cols,
        lag_steps=args.lag_steps,
        max_classes=args.max_classes,
        d_model=args.d_model,
        heads=args.heads,
        table_layers=table_layers,
        graph_layers=graph_layers,
        use_improved=not args.no_improved,
        typed_graph_attention=args.typed_graph_attention,
        typed_task_conditioning=args.typed_task_conditioning,
        row_bridge_attention=args.row_bridge_attention,
    ).to(device)
    if distributed:
        model = DistributedDataParallel(model, device_ids=[local_rank])

    aux_heads: nn.Module | None = None
    raw_aux_heads: nn.Module | None = None
    if args.mechanism_aux_loss_weight > 0 or args.schema_aux_loss_weight > 0:
        num_mechanisms = 36 if args.extended_mechanisms else 24
        aux_heads = StructuralAuxiliaryHeads(args.d_model, num_mechanisms=num_mechanisms).to(device)
        if distributed:
            aux_heads = DistributedDataParallel(aux_heads, device_ids=[local_rank])
        raw_aux_heads = aux_heads.module if distributed else aux_heads

    opt_params = list(model.parameters())
    if aux_heads is not None:
        opt_params.extend(aux_heads.parameters())
    opt = torch.optim.AdamW(opt_params, lr=args.lr, weight_decay=args.weight_decay)
    scheduler = build_scheduler(opt, args.steps, args.warmup_steps, args.min_lr_ratio)
    scaler = torch.amp.GradScaler(
        "cuda",
        enabled=args.amp and device.type == "cuda",
        init_scale=args.amp_init_scale,
        growth_interval=args.amp_growth_interval,
    )
    start_step = 0
    raw_model = model.module if distributed else model
    ema_model: dict[str, torch.Tensor] | None = None
    if args.ema_decay > 0:
        ema_model = {k: v.detach().clone().float() for k, v in raw_model.state_dict().items()}
    validation_history: list[dict[str, object]] = []
    best_validation_score = float("-inf")
    best_checkpoint: str | None = None
    if args.resume is not None:
        checkpoint = torch.load(args.resume, map_location=device, weights_only=False)
        raw_model.load_state_dict(checkpoint["model"])
        if raw_aux_heads is not None and "auxiliary_heads" in checkpoint:
            raw_aux_heads.load_state_dict(checkpoint["auxiliary_heads"])
        opt.load_state_dict(checkpoint["optimizer"])
        if "scheduler" in checkpoint:
            scheduler.load_state_dict(checkpoint["scheduler"])
        if "scaler" in checkpoint:
            scaler.load_state_dict(checkpoint["scaler"])
        start_step = int(checkpoint.get("global_step", 0))
        if "train_rng_state" in checkpoint:
            train_gen.rng.bit_generator.state = checkpoint["train_rng_state"]
        if "torch_rng_state" in checkpoint:
            torch.set_rng_state(checkpoint["torch_rng_state"].cpu())
        if torch.cuda.is_available() and checkpoint.get("cuda_rng_state_all") is not None:
            torch.cuda.set_rng_state_all([state.cpu() for state in checkpoint["cuda_rng_state_all"]])
        checkpoint_manifest = checkpoint.get("manifest", {})
        if isinstance(checkpoint_manifest, dict):
            prior_history = checkpoint_manifest.get("validation_history", [])
            if isinstance(prior_history, list):
                validation_history = [item for item in prior_history if isinstance(item, dict)]
            prior_best = checkpoint_manifest.get("best_validation_score")
            if isinstance(prior_best, (int, float)) and np.isfinite(prior_best):
                best_validation_score = float(prior_best)
            prior_best_checkpoint = checkpoint_manifest.get("best_checkpoint")
            if isinstance(prior_best_checkpoint, str):
                best_checkpoint = prior_best_checkpoint

    iterator = trange(start_step, args.steps, disable=rank != 0)
    model.train()
    train_started = time()
    last_loss = float("nan")
    last_aux_metrics: dict[str, float] = {}
    skipped_nonfinite_steps = 0
    completed_steps = 0
    last_validation_record: dict[str, object] | None = validation_history[-1] if validation_history else None

    def run_validation(global_step: int) -> dict[str, object]:
        nonlocal best_validation_score, best_checkpoint, last_validation_record
        if ema_model is not None:
            with EMAModel(ema_model, raw_model):
                metrics = evaluate_metrics(raw_model, val_gen, device, args.val_batches, args.batch_size)
        else:
            metrics = evaluate_metrics(raw_model, val_gen, device, args.val_batches, args.batch_size)
        score = distributed_finite_mean(metrics["mean_score"], device, distributed)
        metrics["mean_score"] = score
        record: dict[str, object] = {
            "global_step": global_step,
            "optimizer_steps": completed_steps,
            "validation_score": score,
            "validation_metrics": metrics,
            "last_train_loss": last_loss,
            "last_aux_metrics": last_aux_metrics,
            "learning_rate": scheduler.get_last_lr()[0],
            "skipped_nonfinite_steps": skipped_nonfinite_steps,
        }
        improved = bool(np.isfinite(score) and score > best_validation_score)
        if rank == 0:
            validation_history.append(record)
            last_validation_record = record
            args.output_dir.mkdir(parents=True, exist_ok=True)
            if improved:
                best_validation_score = score
                if args.save_best_checkpoint:
                    best_checkpoint = "checkpoint_best.pt"
                record["is_best"] = True
            (args.output_dir / "validation_history.json").write_text(
                json.dumps(json_safe(validation_history), allow_nan=False, indent=2, sort_keys=True) + "\n"
            )
            if args.save_best_checkpoint and improved and best_checkpoint is not None:
                save_checkpoint(
                    args.output_dir / best_checkpoint,
                    raw_model,
                    opt,
                    scheduler,
                    scaler,
                    {
                        "global_step": global_step,
                        "validation_score": score,
                        "validation_metrics": metrics,
                        "best_validation_score": best_validation_score,
                        "best_checkpoint": best_checkpoint,
                        "validation_history": validation_history,
                        "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
                    },
                    global_step,
                    train_gen,
                    raw_aux_heads,
                    ema_state=ema_model,
                )
        return record

    for step in iterator:
        opt.zero_grad(set_to_none=True)
        accum_loss = 0.0
        finite_step = True
        for accum_idx in range(args.grad_accum_steps):
            use_real = real_sampler is not None and float(np.random.random()) < args.real_data_frac
            batch = (real_sampler if use_real else train_gen).batch(args.batch_size).to(device)
            sync_context = nullcontext()
            if distributed and accum_idx < args.grad_accum_steps - 1:
                sync_context = model.no_sync()
            with sync_context:
                ctx = torch.amp.autocast("cuda", enabled=args.amp) if device.type == "cuda" else nullcontext()
                with ctx:
                    if aux_heads is None:
                        logits = logits_from_batch(model, batch)
                        aux_loss = logits.sum() * 0.0
                        aux_metrics = {}
                    else:
                        logits, embedding = model(
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
                            return_embedding=True,
                        )
                        aux_loss, aux_metrics = structural_auxiliary_loss(
                            aux_heads,
                            embedding,
                            batch,
                            mechanism_weight=args.mechanism_aux_loss_weight,
                            schema_weight=args.schema_aux_loss_weight,
                        )
                    loss = (task_loss(logits, batch) + aux_loss) / args.grad_accum_steps
                finite_loss = torch.isfinite(loss).float()
                if distributed:
                    dist.all_reduce(finite_loss, op=dist.ReduceOp.MIN)
                if finite_loss.item() < 1.0:
                    finite_step = False
                    break
                scaler.scale(loss).backward()
                accum_loss += float(loss.detach().cpu()) * args.grad_accum_steps
                if args.context_reconstruction_loss_weight > 0:
                    with ctx:
                        reconstruction_loss, reconstruction_metrics = context_reconstruction_loss(
                            model,
                            batch,
                            weight=args.context_reconstruction_loss_weight,
                        )
                        reconstruction_loss = reconstruction_loss / args.grad_accum_steps
                    finite_reconstruction_loss = torch.isfinite(reconstruction_loss).float()
                    if distributed:
                        dist.all_reduce(finite_reconstruction_loss, op=dist.ReduceOp.MIN)
                    if finite_reconstruction_loss.item() < 1.0:
                        finite_step = False
                        break
                    scaler.scale(reconstruction_loss).backward()
                    accum_loss += float(reconstruction_loss.detach().cpu()) * args.grad_accum_steps
                    aux_metrics = {**aux_metrics, **reconstruction_metrics}
        if not finite_step:
            skipped_nonfinite_steps += 1
            opt.zero_grad(set_to_none=True)
            if rank == 0:
                iterator.set_postfix(loss="nonfinite-skip", world=world)
            continue
        scaler.unscale_(opt)
        finite_grad = finite_gradients(raw_model, raw_aux_heads)
        if distributed:
            dist.all_reduce(finite_grad, op=dist.ReduceOp.MIN)
        if finite_grad.item() < 1.0:
            skipped_nonfinite_steps += 1
            opt.zero_grad(set_to_none=True)
            scaler.update()
            if rank == 0:
                iterator.set_postfix(loss="nonfinite-grad-skip", world=world)
            continue
        nn.utils.clip_grad_norm_(opt_params, 1.0)
        old_scale = scaler.get_scale()
        scaler.step(opt)
        scaler.update()
        if not scaler.is_enabled() or scaler.get_scale() >= old_scale:
            scheduler.step()
        completed_steps += 1
        if ema_model is not None:
            decay = args.ema_decay
            with torch.no_grad():
                for name, param in raw_model.state_dict().items():
                    ema_model[name].mul_(decay).add_(param.detach().float(), alpha=1.0 - decay)
        last_loss = accum_loss / max(1, args.grad_accum_steps)
        last_aux_metrics = aux_metrics
        if rank == 0 and (step + 1) % max(1, args.log_every) == 0:
            iterator.set_postfix(loss=f"{last_loss:.4f}", lr=f"{scheduler.get_last_lr()[0]:.2e}", world=world)
        if args.validate_every > 0 and (step + 1) % args.validate_every == 0:
            validation_record = run_validation(step + 1)
            if rank == 0:
                iterator.set_postfix(
                    loss=f"{last_loss:.4f}",
                    val=f"{validation_record['validation_score']:.4f}",
                    lr=f"{scheduler.get_last_lr()[0]:.2e}",
                    world=world,
                )
        if rank == 0 and args.checkpoint_every > 0 and (step + 1) % args.checkpoint_every == 0:
            args.output_dir.mkdir(parents=True, exist_ok=True)
            interim_manifest = {
                "global_step": step + 1,
                "last_loss": last_loss,
                "optimizer_steps": completed_steps,
                "skipped_nonfinite_steps": skipped_nonfinite_steps,
                "learning_rate": scheduler.get_last_lr()[0],
                "last_validation": last_validation_record,
                "last_aux_metrics": last_aux_metrics,
                "best_validation_score": best_validation_score if np.isfinite(best_validation_score) else None,
                "best_checkpoint": best_checkpoint,
                "validation_history": validation_history,
                "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
            }
            save_checkpoint(
                args.output_dir / f"checkpoint_step_{step + 1}.pt",
                raw_model,
                opt,
                scheduler,
                scaler,
                interim_manifest,
                step + 1,
                train_gen,
                raw_aux_heads,
                ema_state=ema_model,
            )
            if ema_model is not None:
                ema_path = args.output_dir / f"checkpoint_step_{step + 1}_ema.pt"
                ema_payload = {
                    "model": {k: v.cpu().clone() for k, v in ema_model.items()},
                    "manifest": {"global_step": step + 1, "ema_decay": args.ema_decay},
                }
                torch.save(ema_payload, ema_path)

    train_elapsed = max(1e-9, time() - train_started)
    val_metrics = evaluate_metrics(raw_model, val_gen, device, args.val_batches, args.batch_size)
    auc = distributed_finite_mean(val_metrics["mean_score"], device, distributed)
    val_metrics["mean_score"] = auc
    eval_metrics: dict[str, float] = {}
    if rank == 0 and (args.eval_mechanisms or args.eval_robustness or args.eval_latent):
        eval_config = GeneratorConfig(
            context_size=args.context_size,
            root_cols=args.root_cols,
            child_cols=args.child_cols,
            aux_cols=args.aux_cols,
            rows_per_child=args.rows_per_child,
            rows_per_aux=args.rows_per_aux,
            lag_steps=args.lag_steps,
            missing_prob=args.missing_prob,
            link_dropout=args.link_dropout,
            schema_dropout=args.schema_dropout,
            task_type=args.task_type,
            max_classes=args.max_classes,
            binary_quantile_min=args.binary_quantile_min,
            binary_quantile_max=args.binary_quantile_max,
            temporal_row_features=args.temporal_row_features,
            entity_history_frac=args.entity_history_frac,
            entity_drift_scale=args.entity_drift_scale,
            root_shortcut_dropout=args.root_shortcut_dropout,
            feature_heterogeneity=args.feature_heterogeneity,
            history_label_frac=args.history_label_frac,
            relational_bridge_frac=args.relational_bridge_frac,
            feature_diversity=args.feature_diversity,
            feature_correlation=args.feature_correlation,
            extended_mechanisms=args.extended_mechanisms,
        column_permutation=args.column_permutation,
            seed=args.seed + 999,
        )
        if args.eval_mechanisms:
            eval_metrics["mechanisms"] = evaluate_mechanisms(raw_model, eval_config, device, args.eval_batches, args.batch_size)
        if args.eval_robustness:
            eval_metrics["robustness"] = evaluate_robustness(raw_model, eval_config, device, args.eval_batches, args.batch_size)
        if args.eval_latent:
            eval_metrics["latent_structure"] = evaluate_latent_structure(
                raw_model, eval_config, device, max(args.eval_batches, 8), args.batch_size
            )
    if rank == 0:
        eval_summary = eval_metric_summary(eval_metrics)
        final_record: dict[str, object] = {
            "stage": "final",
            "global_step": args.steps,
            "optimizer_steps": completed_steps,
            "validation_score": auc,
            "validation_metrics": val_metrics,
            "last_train_loss": last_loss,
            "last_aux_metrics": last_aux_metrics,
            "learning_rate": scheduler.get_last_lr()[0],
            "skipped_nonfinite_steps": skipped_nonfinite_steps,
        }
        if np.isfinite(auc) and auc > best_validation_score:
            best_validation_score = auc
            final_record["is_best"] = True
            if args.save_best_checkpoint:
                best_checkpoint = "checkpoint_best.pt"
        validation_history.append(final_record)
        print(f"validation_score={auc:.4f} world_size={world} device={device}")
        args.output_dir.mkdir(parents=True, exist_ok=True)
        manifest = {
            "validation_auc": auc,
            "validation_score": auc,
            "validation_metrics": val_metrics,
            "world_size": world,
            "device": str(device),
            "elapsed_sec": time() - started,
            "train_elapsed_sec": train_elapsed,
            "optimizer_steps": completed_steps,
            "global_step": args.steps,
            "resumed_from_step": start_step,
            "effective_batch_size": args.batch_size * world * args.grad_accum_steps,
            "train_tasks_per_sec": completed_steps * args.batch_size * world * args.grad_accum_steps / train_elapsed,
            "train_context_examples_per_sec": completed_steps
            * args.batch_size
            * world
            * args.grad_accum_steps
            * (args.context_size + 1)
            / train_elapsed,
            "last_train_loss": last_loss,
            "last_aux_metrics": last_aux_metrics,
            "skipped_nonfinite_steps": skipped_nonfinite_steps,
            "learning_rate": scheduler.get_last_lr()[0],
            "best_validation_score": best_validation_score if np.isfinite(best_validation_score) else None,
            "best_checkpoint": best_checkpoint,
            "validation_history": validation_history,
            "eval_metrics": eval_metrics,
            "eval_summary": eval_summary,
            "synthetic_mechanism_mean_score": eval_summary.get("mechanisms", {}).get("mean")
            if isinstance(eval_summary.get("mechanisms"), dict)
            else None,
            "robustness_mean_score": eval_summary.get("robustness", {}).get("mean")
            if isinstance(eval_summary.get("robustness"), dict)
            else None,
            "latent_structure_mean_score": eval_summary.get("latent_structure", {}).get("mean")
            if isinstance(eval_summary.get("latent_structure"), dict)
            else None,
            "checkpoint_path": str(args.output_dir / "checkpoint.pt") if args.save_checkpoint else None,
            "best_checkpoint_path": str(args.output_dir / best_checkpoint) if best_checkpoint is not None else None,
            "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
        }
        (args.output_dir / "manifest.json").write_text(
            json.dumps(json_safe(manifest), allow_nan=False, indent=2, sort_keys=True) + "\n"
        )
        (args.output_dir / "validation_history.json").write_text(
            json.dumps(json_safe(validation_history), allow_nan=False, indent=2, sort_keys=True) + "\n"
        )
        if args.save_best_checkpoint and best_checkpoint == "checkpoint_best.pt" and final_record.get("is_best"):
            save_checkpoint(
                args.output_dir / best_checkpoint,
                raw_model,
                opt,
                scheduler,
                scaler,
                manifest,
                args.steps,
                train_gen,
                raw_aux_heads,
                ema_state=ema_model,
            )
        if args.save_checkpoint:
            save_checkpoint(
                args.output_dir / "checkpoint.pt",
                raw_model,
                opt,
                scheduler,
                scaler,
                manifest,
                args.steps,
                train_gen,
                raw_aux_heads,
                ema_state=ema_model,
            )
    if distributed:
        dist.destroy_process_group()


if __name__ == "__main__":
    main()
