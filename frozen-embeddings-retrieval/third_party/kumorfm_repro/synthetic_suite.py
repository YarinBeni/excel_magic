from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
import torch

from .eval import (
    GeneratorConfig,
    evaluate_latent_structure,
    evaluate_mechanisms,
    evaluate_metrics,
    evaluate_robustness,
    make_generator,
)
from .model import RelationalFoundationModel
from .train import json_safe


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run deterministic synthetic scientific validation on a saved checkpoint.")
    p.add_argument("--checkpoint", type=Path, required=True)
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--device", default="cuda")
    p.add_argument("--batches", type=int, default=8)
    p.add_argument("--batch-size", type=int, default=8)
    p.add_argument("--latent-batches", type=int, default=12)
    p.add_argument("--seed", type=int, default=12345)
    p.add_argument("--repeat-check", action="store_true")
    p.add_argument("--task-type", choices=["binary", "regression", "multiclass", "mixed"], default=None)
    return p.parse_args()


def main() -> None:
    args = parse_args()
    device = torch.device(args.device if torch.cuda.is_available() or args.device == "cpu" else "cpu")
    checkpoint = torch.load(args.checkpoint, map_location=device, weights_only=False)
    manifest = checkpoint.get("manifest", {})
    train_args = manifest.get("args", {})
    config = config_from_args(train_args, seed=args.seed, task_type=args.task_type)
    model = build_model(train_args, config, checkpoint, device)

    base = evaluate_metrics(model, make_generator(config, seed_offset=0), device, args.batches, args.batch_size)
    mechanisms = evaluate_mechanisms(model, config, device, args.batches, args.batch_size)
    robustness = evaluate_robustness(model, config, device, args.batches, args.batch_size)
    latent = evaluate_latent_structure(model, config, device, max(args.latent_batches, 8), args.batch_size)
    compositional = {
        "mechanism_7_metapath": mechanisms.get("mechanism_7"),
        "mechanism_8_temporal_composition": mechanisms.get("mechanism_8"),
        "mechanism_9_subgroup": mechanisms.get("mechanism_9"),
        "mechanism_10_multiscale_temporal": mechanisms.get("mechanism_10"),
        "mechanism_12_conditional_aggregation": mechanisms.get("mechanism_12"),
        "mechanism_16_three_hop_metapath": mechanisms.get("mechanism_16"),
        "mechanism_17_variance": mechanisms.get("mechanism_17"),
        "mechanism_18_event_counting": mechanisms.get("mechanism_18"),
        "mechanism_19_recency": mechanisms.get("mechanism_19"),
        "mechanism_20_trend": mechanisms.get("mechanism_20"),
        "mechanism_21_peer_comparison": mechanisms.get("mechanism_21"),
        "mechanism_22_diversity": mechanisms.get("mechanism_22"),
        "mechanism_23_cross_entity": mechanisms.get("mechanism_23"),
        "mean": finite_mean([
            mechanisms.get("mechanism_7"), mechanisms.get("mechanism_8"),
            mechanisms.get("mechanism_9"), mechanisms.get("mechanism_10"),
            mechanisms.get("mechanism_12"), mechanisms.get("mechanism_16"),
            mechanisms.get("mechanism_17"), mechanisms.get("mechanism_18"),
            mechanisms.get("mechanism_19"), mechanisms.get("mechanism_20"),
            mechanisms.get("mechanism_21"), mechanisms.get("mechanism_22"),
            mechanisms.get("mechanism_23"),
        ]),
    }

    repeatability = {}
    if args.repeat_check:
        repeated = evaluate_metrics(model, make_generator(config, seed_offset=0), device, args.batches, args.batch_size)
        repeatability = {
            "base_repeat": repeated,
            "max_abs_delta": max_abs_delta(base, repeated),
            "deterministic": max_abs_delta(base, repeated) <= 1e-8,
        }

    payload = {
        "suite": "synthetic_scientific_validation_v1",
        "checkpoint": str(args.checkpoint),
        "device": str(device),
        "config": config.__dict__,
        "base": base,
        "mechanisms": mechanisms,
        "robustness": robustness,
        "latent_structure": latent,
        "compositional_generalization": compositional,
        "repeatability": repeatability,
        "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    print(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True))


def config_from_args(train_args: dict, *, seed: int, task_type: str | None) -> GeneratorConfig:
    return GeneratorConfig(
        context_size=int(train_args.get("context_size", 8)),
        root_cols=int(train_args.get("root_cols", 8)),
        child_cols=int(train_args.get("child_cols", 8)),
        aux_cols=int(train_args.get("aux_cols", 6)),
        rows_per_child=int(train_args.get("rows_per_child", 4)),
        rows_per_aux=int(train_args.get("rows_per_aux", 3)),
        lag_steps=int(train_args.get("lag_steps", 4)),
        missing_prob=float(train_args.get("missing_prob", 0.05)),
        link_dropout=float(train_args.get("link_dropout", 0.0)),
        schema_dropout=float(train_args.get("schema_dropout", 0.0)),
        task_type=task_type or str(train_args.get("task_type", "mixed")),
        max_classes=int(train_args.get("max_classes", 8)),
        seed=seed,
        binary_quantile_min=float(train_args.get("binary_quantile_min", 0.5)),
        binary_quantile_max=float(train_args.get("binary_quantile_max", 0.5)),
        temporal_row_features=bool(train_args.get("temporal_row_features", True)),
        entity_history_frac=float(train_args.get("entity_history_frac", 0.0)),
        entity_drift_scale=float(train_args.get("entity_drift_scale", 0.3)),
        root_shortcut_dropout=float(train_args.get("root_shortcut_dropout", 0.0)),
        feature_heterogeneity=float(train_args.get("feature_heterogeneity", 0.0)),
        history_label_frac=float(train_args.get("history_label_frac", 0.0)),
        relational_bridge_frac=float(train_args.get("relational_bridge_frac", 0.0)),
    )


def build_model(train_args: dict, config: GeneratorConfig, checkpoint: dict, device: torch.device) -> RelationalFoundationModel:
    model = RelationalFoundationModel(
        root_cols=config.root_cols,
        child_cols=config.child_cols,
        aux_cols=config.aux_cols,
        lag_steps=config.lag_steps,
        max_classes=config.max_classes,
        d_model=int(train_args.get("d_model", 256)),
        heads=int(train_args.get("heads", 8)),
        table_layers=int(train_args.get("table_layers") or train_args.get("layers", 3)),
        graph_layers=int(train_args.get("graph_layers") or train_args.get("layers", 3)),
        use_improved=not bool(train_args.get("no_improved")),
        typed_graph_attention=bool(train_args.get("typed_graph_attention", False)),
        typed_task_conditioning=bool(train_args.get("typed_task_conditioning", False)),
        row_bridge_attention=bool(train_args.get("row_bridge_attention", False)),
    ).to(device)
    model.load_state_dict(checkpoint["model"])
    model.eval()
    return model
    model.load_state_dict(checkpoint["model"])
    model.eval()
    return model


def finite_mean(values: list[float | None]) -> float:
    finite = [float(v) for v in values if v is not None and np.isfinite(float(v))]
    return float(np.mean(finite)) if finite else float("nan")


def max_abs_delta(left: dict[str, float], right: dict[str, float]) -> float:
    deltas = []
    for key in sorted(set(left) & set(right)):
        a = left[key]
        b = right[key]
        if isinstance(a, float) and isinstance(b, float) and np.isfinite(a) and np.isfinite(b):
            deltas.append(abs(a - b))
    return float(max(deltas)) if deltas else float("nan")


if __name__ == "__main__":
    main()
