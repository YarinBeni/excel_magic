from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from time import time
from typing import Any

import numpy as np
import torch

from .icl_suite import run_suite_from_args, target_key_for_preset
from .targets import compare_manifest_to_target
from .train import json_safe


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run fixed no-finetuning ICL screening across multiple checkpoints.")
    p.add_argument("--checkpoint", action="append", type=Path, default=[], help="Checkpoint to evaluate. Can be passed more than once.")
    p.add_argument("--checkpoint-dir", type=Path, default=None, help="Directory containing checkpoint*.pt files.")
    p.add_argument("--checkpoint-glob", default="checkpoint*.pt", help="Glob used with --checkpoint-dir.")
    p.add_argument("--output-dir", type=Path, required=True)
    p.add_argument("--data-dir", type=Path, default=Path("data/relbench_exports"))
    p.add_argument("--device", default="cuda")
    p.add_argument("--context-size", type=int, default=16)
    p.add_argument("--batch-size", type=int, default=16)
    p.add_argument("--rows-per-child", type=int, default=8)
    p.add_argument("--rows-per-aux", type=int, default=6)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--child-cols", type=int, default=8)
    p.add_argument("--aux-cols", type=int, default=6)
    p.add_argument("--max-eval-samples", type=int, default=128)
    p.add_argument("--eval-min-target-classes", type=int, default=1)
    p.add_argument("--eval-diversity-scan-limit", type=int, default=0)
    p.add_argument("--train-frac", type=float, default=0.7)
    p.add_argument("--context-mode", choices=["recent", "entity", "entity_mixed"], default="entity_mixed")
    p.add_argument("--local-context-frac", type=float, default=0.5)
    p.add_argument("--root-baseline-max-train", type=int, default=20000)
    p.add_argument("--indexed-table-lookup", action="store_true")
    p.add_argument("--skip-root-baseline", action="store_true")
    p.add_argument("--context-ablation", action="append", choices=["zero_targets", "shuffle_targets"], default=[])
    p.add_argument("--task", action="append", default=[])
    p.add_argument(
        "--preset",
        choices=["relbenchv1_classification", "relbenchv1_regression", "relbenchv2_classification", "relbenchv2_regression"],
        default="relbenchv1_classification",
    )
    p.add_argument("--max-tasks", type=int, default=2)
    p.add_argument("--metric-path", default=None, help="Metric path used for ranking and target comparison.")
    p.add_argument("--rank-direction", choices=["higher", "lower"], default="higher")
    p.add_argument("--skip-existing", action="store_true", help="Reuse per-checkpoint icl_suite_summary.json files when present.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    started = time()
    checkpoints = discover_checkpoints(args)
    if not checkpoints:
        raise SystemExit("No checkpoints found. Pass --checkpoint or --checkpoint-dir.")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    metric_path = args.metric_path or default_metric_path(args.preset)
    target_key = target_key_for_preset(args.preset)
    entries: list[dict[str, Any]] = []

    for checkpoint in checkpoints:
        label = checkpoint_label(checkpoint)
        checkpoint_output = args.output_dir / label
        summary_path = checkpoint_output / "icl_suite_summary.json"
        if args.skip_existing and summary_path.exists():
            payload = json.loads(summary_path.read_text())
        else:
            suite_args = argparse.Namespace(
                checkpoint=checkpoint,
                output_dir=checkpoint_output,
                data_dir=args.data_dir,
                device=args.device,
                context_size=args.context_size,
                batch_size=args.batch_size,
                rows_per_child=args.rows_per_child,
                rows_per_aux=args.rows_per_aux,
                root_cols=args.root_cols,
                child_cols=args.child_cols,
                aux_cols=args.aux_cols,
                max_eval_samples=args.max_eval_samples,
                eval_min_target_classes=args.eval_min_target_classes,
                eval_diversity_scan_limit=args.eval_diversity_scan_limit,
                train_frac=args.train_frac,
                context_mode=args.context_mode,
                local_context_frac=args.local_context_frac,
                root_baseline_max_train=args.root_baseline_max_train,
                indexed_table_lookup=args.indexed_table_lookup,
                skip_root_baseline=args.skip_root_baseline,
                context_ablation=args.context_ablation,
                task=args.task,
                preset=args.preset,
                max_tasks=args.max_tasks,
            )
            payload = run_suite_from_args(suite_args)
        metric = lookup_path(payload, metric_path)
        comparison = None
        if target_key is not None:
            comparison = compare_manifest_to_target(summary_path, target_key, metric_path=metric_path)
        entries.append(
            {
                "checkpoint": str(checkpoint),
                "label": label,
                "summary_path": str(summary_path),
                "metric_path": metric_path,
                "metric": metric,
                "metric_normalized": normalize_metric(metric, metric_path),
                "target_comparison": comparison,
                "summary": payload.get("summary"),
                "aggregate": payload.get("aggregate"),
                "checkpoint_info": checkpoint_info(checkpoint),
            }
        )
        if torch.cuda.is_available():
            torch.cuda.empty_cache()

    ranked = sorted(entries, key=lambda item: rank_key(item, args.rank_direction))
    best = ranked[0] if ranked else None
    payload = {
        "sweep": "transfer_checkpoint_sweep_v1",
        "evidence_type": "no_finetuning_icl_model_selection_diagnostic",
        "elapsed_sec": time() - started,
        "metric_path": metric_path,
        "rank_direction": args.rank_direction,
        "best": best,
        "checkpoints": ranked,
        "args": {k: str(v) if isinstance(v, Path) else [str(x) for x in v] if k == "checkpoint" else v for k, v in vars(args).items()},
    }
    output_path = args.output_dir / "transfer_sweep_summary.json"
    output_path.write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    print(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True))


def discover_checkpoints(args: argparse.Namespace) -> list[Path]:
    checkpoints = list(args.checkpoint)
    if args.checkpoint_dir is not None:
        checkpoints.extend(sorted(args.checkpoint_dir.glob(args.checkpoint_glob), key=checkpoint_sort_key))
    seen: set[Path] = set()
    unique: list[Path] = []
    for checkpoint in checkpoints:
        checkpoint = checkpoint.expanduser()
        if checkpoint in seen or not checkpoint.exists():
            continue
        seen.add(checkpoint)
        unique.append(checkpoint)
    return unique


def checkpoint_sort_key(path: Path) -> tuple[int, str]:
    name = path.name
    match = re.search(r"checkpoint_step_(\d+)\.pt$", name)
    if match:
        return (int(match.group(1)), name)
    if name == "checkpoint_best.pt":
        return (10**12 - 1, name)
    if name == "checkpoint.pt":
        return (10**12, name)
    return (10**12 + 1, name)


def checkpoint_label(path: Path) -> str:
    if path.name.endswith(".pt"):
        return path.stem
    return path.name


def checkpoint_info(path: Path) -> dict[str, Any]:
    try:
        checkpoint = torch.load(path, map_location="cpu", weights_only=False)
    except Exception as exc:
        return {"error": repr(exc)}
    manifest = checkpoint.get("manifest", {})
    if not isinstance(manifest, dict):
        return {}
    args = manifest.get("args", {}) if isinstance(manifest.get("args"), dict) else {}
    return {
        "global_step": checkpoint.get("global_step", manifest.get("global_step")),
        "optimizer_steps": manifest.get("optimizer_steps"),
        "validation_score": manifest.get("validation_score"),
        "best_validation_score": manifest.get("best_validation_score"),
        "last_train_loss": manifest.get("last_train_loss"),
        "args": {
            key: args.get(key)
            for key in (
                "steps",
                "context_size",
                "d_model",
                "layers",
                "entity_history_frac",
                "root_shortcut_dropout",
                "history_label_frac",
                "feature_heterogeneity",
                "relational_bridge_frac",
            )
            if key in args
        },
    }


def default_metric_path(preset: str) -> str:
    if "classification" in preset:
        return "aggregate.classification.avg_auroc"
    if "regression" in preset:
        return "aggregate.regression.avg_mae"
    return "summary.avg_auroc"


def lookup_path(payload: dict[str, Any], path: str) -> Any:
    cur: Any = payload
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur


def normalize_metric(metric: Any, metric_path: str) -> float | None:
    if metric is None:
        return None
    value = float(metric)
    if metric_path.endswith("avg_auroc") and 0.0 <= value <= 1.0:
        return value * 100.0
    return value


def rank_key(item: dict[str, Any], direction: str) -> float:
    metric = item.get("metric")
    if metric is None:
        return float("inf")
    value = float(metric)
    return -value if direction == "higher" else value


if __name__ == "__main__":
    main()
