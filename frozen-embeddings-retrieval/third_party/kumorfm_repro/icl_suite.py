from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from time import time

import numpy as np
import torch

from .icl_eval import evaluate_icl, icl_summary, load_model_from_checkpoint
from .relational_io import RelationalDataset
from .targets import PAPER_TARGETS
from .train import json_safe

TASK_DIR_MAP = {
    "rel-f1:driver-dnf": "rel-f1_driver-dnf_train",
    "rel-f1:driver-top3": "rel-f1_driver-top3_train",
    "rel-avito:user-clicks": "rel-avito_user-clicks_train",
    "rel-avito:user-visits": "rel-avito_user-visits_train",
    "rel-event:user-repeat": "rel-event_user-repeat_train",
    "rel-event:user-ignore": "rel-event_user-ignore_train",
    "rel-trial:study-outcome": "rel-trial_study-outcome_train",
    "rel-amazon:user-churn": "rel-amazon_user-churn_train",
    "rel-amazon:item-churn": "rel-amazon_item-churn_train",
    "rel-stack:user-engagement": "rel-stack_user-engagement_train",
    "rel-stack:user-badge": "rel-stack_user-badge_train",
    "rel-hm:user-churn": "rel-hm_user-churn_train",
}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run in-context evaluation of a pre-trained model on RelBench tasks.")
    p.add_argument("--checkpoint", type=Path, required=True)
    p.add_argument("--output-dir", type=Path, default=Path("runs/icl_eval"))
    p.add_argument("--data-dir", type=Path, default=Path("data/relbench_exports"))
    p.add_argument("--device", default="cuda")
    p.add_argument("--context-size", type=int, default=32)
    p.add_argument("--batch-size", type=int, default=32)
    p.add_argument("--rows-per-child", type=int, default=12)
    p.add_argument("--rows-per-aux", type=int, default=8)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--child-cols", type=int, default=8)
    p.add_argument("--aux-cols", type=int, default=6)
    p.add_argument("--max-eval-samples", type=int, default=2048)
    p.add_argument("--eval-min-target-classes", type=int, default=1)
    p.add_argument("--eval-diversity-scan-limit", type=int, default=0)
    p.add_argument("--train-frac", type=float, default=0.7)
    p.add_argument("--context-mode", choices=["recent", "entity", "entity_mixed"], default="recent")
    p.add_argument("--local-context-frac", type=float, default=0.5)
    p.add_argument("--root-baseline-max-train", type=int, default=0)
    p.add_argument("--indexed-table-lookup", action="store_true")
    p.add_argument("--skip-root-baseline", action="store_true")
    p.add_argument("--context-ablation", action="append", choices=["zero_targets", "shuffle_targets"], default=[])
    p.add_argument("--task", action="append", default=[])
    p.add_argument("--preset", choices=["relbenchv1_classification", "relbenchv1_regression", "relbenchv2_classification", "relbenchv2_regression"], default="relbenchv1_classification")
    p.add_argument("--max-tasks", type=int, default=0)
    return p.parse_args()


TASK_LISTS = {
    "relbenchv1_classification": [
        "rel-f1:driver-dnf", "rel-f1:driver-top3",
        "rel-avito:user-clicks", "rel-avito:user-visits",
        "rel-event:user-repeat", "rel-event:user-ignore",
        "rel-trial:study-outcome",
        "rel-amazon:user-churn", "rel-amazon:item-churn",
        "rel-stack:user-engagement", "rel-stack:user-badge",
        "rel-hm:user-churn",
    ],
    "relbenchv1_regression": [
        "rel-f1:driver-position",
        "rel-avito:ad-ctr",
        "rel-event:user-attendance",
        "rel-trial:study-adverse", "rel-trial:site-success",
        "rel-amazon:user-ltv", "rel-amazon:item-ltv",
        "rel-stack:post-votes",
        "rel-hm:item-sales",
    ],
    "relbenchv2_classification": [
        "rel-mimic:patient-iculengthofstay",
        "rel-ratebeer:beer-churn", "rel-ratebeer:user-churn",
        "rel-ratebeer:brewer-dormant",
        "rel-arxiv:paper-citation",
    ],
    "relbenchv2_regression": [
        "rel-ratebeer:user-count",
        "rel-arxiv:author-publication",
    ],
}


def main() -> None:
    args = parse_args()
    payload = run_suite_from_args(args)
    print(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True))


def run_suite_from_args(args: argparse.Namespace) -> dict:
    device = torch.device(args.device if torch.cuda.is_available() or args.device == "cpu" else "cpu")
    started = time()

    tasks = list(args.task) if args.task else TASK_LISTS.get(args.preset, TASK_LISTS["relbenchv1_classification"])
    if args.max_tasks > 0:
        tasks = tasks[: args.max_tasks]

    model = load_model_from_checkpoint(args.checkpoint, device)
    results = []

    for task_spec in tasks:
        dir_name = TASK_DIR_MAP.get(task_spec)
        if dir_name is None:
            dataset_part, task_part = task_spec.split(":", 1)
            dir_name = f"{dataset_part}_{task_part}_train"
        dataset_dir = args.data_dir / dir_name
        entry = {"task": task_spec, "dataset_dir": str(dataset_dir)}
        if not (dataset_dir / "metadata.json").exists():
            entry["ok"] = False
            entry["error"] = "dataset not found"
            results.append(entry)
            continue
        try:
            dataset = RelationalDataset.load(dataset_dir)
            task_args = argparse.Namespace(
                dataset_dir=str(dataset_dir),
                context_size=args.context_size,
                batch_size=args.batch_size,
                rows_per_child=args.rows_per_child,
                rows_per_aux=args.rows_per_aux,
                root_cols=args.root_cols,
                child_cols=args.child_cols,
                aux_cols=args.aux_cols,
                lag_steps=getattr(args, "lag_steps", 4),
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
            )
            task_started = time()
            metrics = evaluate_icl(model, dataset, task_args, device)
            entry["ok"] = "error" not in metrics
            entry.update(metrics)
            entry["elapsed_sec"] = time() - task_started
        except Exception as exc:
            entry["ok"] = False
            entry["error"] = repr(exc)
        results.append(entry)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    summary_info = icl_summary(results)
    payload = {
        "checkpoint": str(args.checkpoint),
        "device": str(device),
        "preset": args.preset,
        "elapsed_sec": time() - started,
        "tasks": results,
        "summary": summary_info,
        "aggregate": aggregate_from_icl_results(results),
        "target_key": target_key_for_preset(args.preset),
        "evidence_type": "no_finetuning_icl",
        "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
    }
    output_path = args.output_dir / "icl_suite_summary.json"
    output_path.write_text(json.dumps(json_safe(payload), allow_nan=False, indent=2, sort_keys=True) + "\n")
    return payload

def aggregate_from_icl_results(results: list[dict]) -> dict:
    ok_tasks = [task for task in results if task.get("ok", False)]
    aurocs = [float(task["binary_auroc"]) for task in ok_tasks if "binary_auroc" in task and np.isfinite(task["binary_auroc"])]
    maes = [float(task["regression_mae"]) for task in ok_tasks if "regression_mae" in task and np.isfinite(task["regression_mae"])]
    mrrs = [float(task["multiclass_mrr"]) for task in ok_tasks if "multiclass_mrr" in task and np.isfinite(task["multiclass_mrr"])]
    root_aurocs = [float(task["root_only_auroc"]) for task in ok_tasks if "root_only_auroc" in task and np.isfinite(task["root_only_auroc"])]
    ablation_summary: dict[str, dict[str, float | int | None]] = {}
    for ablation_name in ("zero_targets", "shuffle_targets"):
        auroc_values = [
            float(task[f"ablation_{ablation_name}_binary_auroc"])
            for task in ok_tasks
            if f"ablation_{ablation_name}_binary_auroc" in task and np.isfinite(task[f"ablation_{ablation_name}_binary_auroc"])
        ]
        auroc_deltas = [
            float(task[f"ablation_{ablation_name}_binary_auroc_delta"])
            for task in ok_tasks
            if f"ablation_{ablation_name}_binary_auroc_delta" in task
            and np.isfinite(task[f"ablation_{ablation_name}_binary_auroc_delta"])
        ]
        mae_deltas = [
            float(task[f"ablation_{ablation_name}_regression_mae_delta"])
            for task in ok_tasks
            if f"ablation_{ablation_name}_regression_mae_delta" in task
            and np.isfinite(task[f"ablation_{ablation_name}_regression_mae_delta"])
        ]
        mrr_deltas = [
            float(task[f"ablation_{ablation_name}_multiclass_mrr_delta"])
            for task in ok_tasks
            if f"ablation_{ablation_name}_multiclass_mrr_delta" in task
            and np.isfinite(task[f"ablation_{ablation_name}_multiclass_mrr_delta"])
        ]
        if auroc_values or auroc_deltas or mae_deltas or mrr_deltas:
            ablation_summary[ablation_name] = {
                "avg_binary_auroc": float(np.mean(auroc_values)) if auroc_values else None,
                "avg_binary_auroc_delta": float(np.mean(auroc_deltas)) if auroc_deltas else None,
                "avg_regression_mae_delta": float(np.mean(mae_deltas)) if mae_deltas else None,
                "avg_multiclass_mrr_delta": float(np.mean(mrr_deltas)) if mrr_deltas else None,
                "num_binary_tasks": len(auroc_values),
            }
    return {
        "num_tasks": len(results),
        "num_ok": len(ok_tasks),
        "classification": {
            "avg_auroc": float(np.mean(aurocs)) if aurocs else None,
            "avg_auroc_points": float(np.mean(aurocs) * 100.0) if aurocs else None,
            "num_tasks": len(aurocs),
            "root_only_avg_auroc": float(np.mean(root_aurocs)) if root_aurocs else None,
        },
        "regression": {
            "avg_mae": float(np.mean(maes)) if maes else None,
            "num_tasks": len(maes),
        },
        "multiclass": {
            "avg_mrr": float(np.mean(mrrs)) if mrrs else None,
            "num_tasks": len(mrrs),
        },
        "context_ablation": ablation_summary,
    }


def target_key_for_preset(preset: str) -> str | None:
    if preset == "relbenchv1_classification":
        return "relbenchv1_classification_avg_auroc"
    if preset == "relbenchv1_regression":
        return "relbenchv1_regression_normalized_mae"
    if preset == "relbenchv2_classification":
        return "relbenchv2_classification_avg_auroc"
    if preset == "relbenchv2_regression":
        return "relbenchv2_regression_normalized_mae"
    return None


if __name__ == "__main__":
    main()
