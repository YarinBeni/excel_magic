from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from time import time


TASK_PRESETS: dict[str, list[str]] = {
    "smoke": ["rel-f1:driver-dnf", "rel-f1:driver-circuit-compete"],
    "relbenchv1_classification": [
        "rel-f1:driver-dnf",
        "rel-f1:driver-top3",
        "rel-avito:user-clicks",
        "rel-avito:user-visits",
        "rel-event:user-repeat",
        "rel-event:user-ignore",
        "rel-trial:study-outcome",
        "rel-amazon:user-churn",
        "rel-amazon:item-churn",
        "rel-stack:user-engagement",
        "rel-stack:user-badge",
        "rel-hm:user-churn",
    ],
    "relbenchv1_regression": [
        "rel-f1:driver-position",
        "rel-avito:ad-ctr",
        "rel-event:user-attendance",
        "rel-trial:study-adverse",
        "rel-trial:site-success",
        "rel-amazon:user-ltv",
        "rel-amazon:item-ltv",
        "rel-stack:post-votes",
        "rel-hm:item-sales",
    ],
    "relbenchv2_classification": [
        "rel-mimic:patient-iculengthofstay",
        "rel-ratebeer:beer-churn",
        "rel-ratebeer:user-churn",
        "rel-ratebeer:brewer-dormant",
        "rel-arxiv:paper-citation",
    ],
    "relbenchv2_regression": [
        "rel-ratebeer:user-count",
        "rel-arxiv:author-publication",
    ],
    "salt_multiclass": [
        "rel-salt:item-plant",
        "rel-salt:item-shippoint",
        "rel-salt:item-incoterms",
        "rel-salt:sales-office",
        "rel-salt:sales-group",
        "rel-salt:sales-payterms",
        "rel-salt:sales-shipcond",
        "rel-salt:sales-incoterms",
    ],
}


PRESET_TARGET_KEYS: dict[str, str] = {
    "relbenchv1_classification": "relbenchv1_classification_avg_auroc",
    "relbenchv1_regression": "relbenchv1_regression_normalized_mae",
    "relbenchv2_classification": "relbenchv2_classification_avg_auroc",
    "relbenchv2_regression": "relbenchv2_regression_normalized_mae",
    "salt_multiclass": "salt_multiclass_mrr_icl",
}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run a small official RelBench export/train/eval suite.")
    p.add_argument("--output-dir", type=Path, default=Path("runs/relbench_suite"))
    p.add_argument("--data-dir", type=Path, default=Path("data/relbench_exports"))
    p.add_argument("--task", action="append", default=[], help="Task spec as dataset:task, for example rel-f1:driver-dnf.")
    p.add_argument("--preset", choices=sorted(TASK_PRESETS), help="Run a predefined paper/smoke task group.")
    p.add_argument("--list-presets", action="store_true", help="Print predefined task groups and exit.")
    p.add_argument("--task-offset", type=int, default=0, help="Skip this many selected tasks before applying --max-tasks.")
    p.add_argument("--max-tasks", type=int, default=0, help="Limit selected tasks for staged/debug runs. 0 means all selected tasks.")
    p.add_argument("--split", default="train", choices=("train", "val", "test"))
    p.add_argument("--download", action="store_true")
    p.add_argument("--max-rows", type=int, default=200_000, help="Maximum rows to keep per exported source table.")
    p.add_argument("--steps", type=int, default=12)
    p.add_argument("--batch-size", type=int, default=8)
    p.add_argument("--context-size", type=int, default=16)
    p.add_argument("--rows-per-child", type=int, default=8)
    p.add_argument("--rows-per-aux", type=int, default=4)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--child-cols", type=int, default=8)
    p.add_argument("--aux-cols", type=int, default=4)
    p.add_argument("--d-model", type=int, default=256)
    p.add_argument("--heads", type=int, default=8)
    p.add_argument("--layers", type=int, default=3)
    p.add_argument("--grad-accum-steps", type=int, default=1)
    p.add_argument("--warmup-steps", type=int, default=0)
    p.add_argument("--min-lr-ratio", type=float, default=0.1)
    p.add_argument("--eval-batches", type=int, default=2)
    p.add_argument("--eval-min-rows", type=int, default=64)
    p.add_argument("--eval-max-batches", type=int, default=128)
    p.add_argument("--balanced-batches", action="store_true", help="Prefer train windows containing multiple labels for classification tasks.")
    p.add_argument("--nproc-per-node", type=int, default=2)
    p.add_argument("--amp", action="store_true")
    p.add_argument("--timeout-sec", type=int, default=600)
    p.add_argument("--continue-on-error", action="store_true")
    p.add_argument("--skip-existing", action="store_true", help="Reuse completed per-task artifacts instead of rerunning export/validation/training.")
    p.add_argument("--dry-run", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if args.list_presets:
        print(json.dumps(preset_payload(), indent=2, sort_keys=True), flush=True)
        return
    tasks = resolve_tasks(args)
    impl_dir = Path(__file__).resolve().parents[1]
    args.output_dir.mkdir(parents=True, exist_ok=True)
    args.data_dir.mkdir(parents=True, exist_ok=True)

    summaries = []
    started = time()
    for spec in tasks:
        dataset, task = parse_task_spec(spec)
        slug = f"{dataset}_{task}_{args.split}".replace("/", "_")
        dataset_dir = args.data_dir / slug
        task_run_dir = args.output_dir / slug
        task_run_dir.mkdir(parents=True, exist_ok=True)
        export_json = task_run_dir / "export.json"
        validation_json = task_run_dir / "validation.json"
        batch_json = task_run_dir / "batch.json"
        train_dir = task_run_dir / "train"
        entry = {"dataset": dataset, "task": task, "split": args.split, "dataset_dir": str(dataset_dir), "run_dir": str(task_run_dir)}
        try:
            entry["planned_commands"] = []
            if args.skip_existing and not args.dry_run and task_artifacts_complete(export_json, validation_json, batch_json, train_dir / "manifest.json"):
                entry.update(read_task_outputs(export_json, validation_json, batch_json, train_dir / "manifest.json"))
                entry["ok"] = True
                entry["skipped_existing"] = True
                summaries.append(entry)
                continue
            export_cmd = [
                sys.executable,
                "-m",
                "kumorfm_repro.benchmark_adapter",
                "relbench-export",
                "--dataset",
                dataset,
                "--task",
                task,
                "--split",
                args.split,
                "--output-dir",
                str(dataset_dir),
                "--output",
                str(export_json),
                "--max-rows",
                str(args.max_rows),
                *((("--download",) if args.download else ())),
            ]
            entry["planned_commands"].append(format_cmd(export_cmd))
            run_cmd(export_cmd, cwd=impl_dir, timeout=args.timeout_sec, dry_run=args.dry_run)
            validate_cmd = [
                sys.executable,
                "-m",
                "kumorfm_repro.benchmark_adapter",
                "validate",
                "--dataset-dir",
                str(dataset_dir),
                "--context-size",
                str(args.context_size),
                "--output",
                str(validation_json),
            ]
            entry["planned_commands"].append(format_cmd(validate_cmd))
            run_cmd(validate_cmd, cwd=impl_dir, timeout=args.timeout_sec, dry_run=args.dry_run)
            batch_cmd = [
                sys.executable,
                "-m",
                "kumorfm_repro.benchmark_adapter",
                "batch-smoke",
                "--dataset-dir",
                str(dataset_dir),
                "--context-size",
                str(args.context_size),
                "--batch-size",
                "4",
                "--rows-per-child",
                str(args.rows_per_child),
                "--rows-per-aux",
                str(args.rows_per_aux),
                "--root-cols",
                str(args.root_cols),
                "--child-cols",
                str(args.child_cols),
                "--aux-cols",
                str(args.aux_cols),
                "--output",
                str(batch_json),
            ]
            entry["planned_commands"].append(format_cmd(batch_cmd))
            run_cmd(batch_cmd, cwd=impl_dir, timeout=args.timeout_sec, dry_run=args.dry_run)
            train_cmd = [
                "torchrun",
                "--standalone",
                f"--nproc_per_node={args.nproc_per_node}",
                "-m",
                "kumorfm_repro.benchmark_adapter",
                "train-model",
                "--dataset-dir",
                str(dataset_dir),
                "--output-dir",
                str(train_dir),
                "--steps",
                str(args.steps),
                "--batch-size",
                str(args.batch_size),
                "--context-size",
                str(args.context_size),
                "--rows-per-child",
                str(args.rows_per_child),
                "--rows-per-aux",
                str(args.rows_per_aux),
                "--root-cols",
                str(args.root_cols),
                "--child-cols",
                str(args.child_cols),
                "--aux-cols",
                str(args.aux_cols),
                "--d-model",
                str(args.d_model),
                "--heads",
                str(args.heads),
                "--layers",
                str(args.layers),
                "--grad-accum-steps",
                str(args.grad_accum_steps),
                "--warmup-steps",
                str(args.warmup_steps),
                "--min-lr-ratio",
                str(args.min_lr_ratio),
                "--eval-batches",
                str(args.eval_batches),
                "--eval-min-rows",
                str(args.eval_min_rows),
                "--eval-max-batches",
                str(args.eval_max_batches),
                "--save-checkpoint",
            ]
            if args.amp:
                train_cmd.append("--amp")
            if args.balanced_batches:
                train_cmd.append("--balanced-batches")
            entry["planned_commands"].append(format_cmd(train_cmd))
            run_cmd(train_cmd, cwd=impl_dir, timeout=args.timeout_sec, dry_run=args.dry_run)
            if args.dry_run:
                entry["planned"] = True
                entry["ok"] = None
            else:
                entry.update(read_task_outputs(export_json, validation_json, batch_json, train_dir / "manifest.json"))
                entry["ok"] = True
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError) as exc:
            entry["ok"] = False
            entry["error"] = repr(exc)
            if not args.continue_on_error:
                summaries.append(entry)
                write_summary(args, summaries, started)
                raise
        summaries.append(entry)

    write_summary(args, summaries, started)


def preset_payload() -> dict:
    return {
        name: {
            "tasks": tasks,
            "num_tasks": len(tasks),
            "target_key": PRESET_TARGET_KEYS.get(name),
        }
        for name, tasks in TASK_PRESETS.items()
    }


def resolve_tasks(args: argparse.Namespace) -> list[str]:
    if args.task and args.preset:
        raise ValueError("Use either --task or --preset, not both.")
    tasks = list(args.task) if args.task else list(TASK_PRESETS[args.preset or "smoke"])
    if args.task_offset < 0:
        raise ValueError("--task-offset must be non-negative.")
    tasks = tasks[args.task_offset :]
    if args.max_tasks:
        if args.max_tasks < 0:
            raise ValueError("--max-tasks must be non-negative.")
        tasks = tasks[: args.max_tasks]
    return tasks


def parse_task_spec(spec: str) -> tuple[str, str]:
    if ":" not in spec:
        raise ValueError(f"Task spec must be dataset:task, got {spec!r}")
    dataset, task = spec.split(":", 1)
    if not dataset or not task:
        raise ValueError(f"Task spec must be dataset:task, got {spec!r}")
    return dataset, task


def format_cmd(cmd: list[str]) -> str:
    return " ".join(cmd)


def run_cmd(cmd: list[str], cwd: Path, timeout: int, dry_run: bool) -> None:
    print(format_cmd(cmd), flush=True)
    if dry_run:
        return
    subprocess.run(cmd, cwd=cwd, check=True, timeout=timeout)


def read_json(path: Path) -> dict:
    return json.loads(path.read_text()) if path.exists() else {}


def task_artifacts_complete(export_json: Path, validation_json: Path, batch_json: Path, manifest_json: Path) -> bool:
    if not all(path.exists() for path in (export_json, validation_json, batch_json, manifest_json)):
        return False
    export = read_json(export_json)
    validation = read_json(validation_json)
    batch = read_json(batch_json)
    manifest = read_json(manifest_json)
    if not export.get("ok"):
        return False
    if not validation.get("validation", {}).get("ok"):
        return False
    if not batch.get("batch"):
        return False
    if not manifest.get("model_metrics"):
        return False
    return int(manifest.get("world_size") or 0) > 0 and int(manifest.get("global_step") or 0) > 0


def read_task_outputs(export_json: Path, validation_json: Path, batch_json: Path, manifest_json: Path) -> dict:
    export = read_json(export_json)
    validation = read_json(validation_json)
    batch = read_json(batch_json)
    manifest = read_json(manifest_json)
    return {
        "export": {
            "rows": export.get("rows"),
            "relbench_task_type": export.get("relbench_task_type"),
            "link_export": export.get("link_export"),
            "selected_tables": export.get("selected_tables"),
        },
        "validation_ok": validation.get("validation", {}).get("ok"),
        "batch_shapes": batch.get("batch"),
        "world_size": manifest.get("world_size"),
        "last_train_loss": manifest.get("last_train_loss"),
        "skipped_nonfinite_steps": manifest.get("skipped_nonfinite_steps"),
        "effective_batch_size": manifest.get("effective_batch_size"),
        "train_tasks_per_sec": manifest.get("train_tasks_per_sec"),
        "train_context_examples_per_sec": manifest.get("train_context_examples_per_sec"),
        "learning_rate": manifest.get("learning_rate"),
        "optimizer_steps": manifest.get("optimizer_steps"),
        "global_step": manifest.get("global_step"),
        "resumed_from_step": manifest.get("resumed_from_step"),
        "model_metrics": manifest.get("model_metrics"),
        "flattened_baseline": manifest.get("flattened_baseline"),
        "train_manifest": str(manifest_json),
    }


def write_summary(args: argparse.Namespace, summaries: list[dict], started: float) -> None:
    payload = {
        "elapsed_sec": time() - started,
        "tasks": summaries,
        "aggregate": aggregate_metrics(summaries),
        "target_key": PRESET_TARGET_KEYS.get(args.preset) if args.preset else None,
        "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()},
    }
    summary_path = args.output_dir / "suite_summary.json"
    summary_path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    print(json.dumps(payload, indent=2, sort_keys=True), flush=True)


def aggregate_metrics(summaries: list[dict]) -> dict:
    ok_tasks = [task for task in summaries if task.get("ok")]
    classification = []
    regression = []
    multiclass = []
    link_mrr = []
    link_hit = []
    link_recall = []
    task_type_counts: dict[str, int] = {}
    throughput = []
    context_throughput = []
    for task in ok_tasks:
        relbench_type = task.get("export", {}).get("relbench_task_type", "unknown")
        task_type_counts[relbench_type] = task_type_counts.get(relbench_type, 0) + 1
        metrics = task.get("model_metrics") or {}
        if "binary_auroc" in metrics:
            classification.append(float(metrics["binary_auroc"]))
        if "regression_mae" in metrics:
            regression.append(float(metrics["regression_mae"]))
        if "multiclass_mrr" in metrics:
            multiclass.append(float(metrics["multiclass_mrr"]))
        if "link_mrr_at_10" in metrics:
            link_mrr.append(float(metrics["link_mrr_at_10"]))
        if "link_hit_at_10" in metrics:
            link_hit.append(float(metrics["link_hit_at_10"]))
        if "link_recall_at_10" in metrics:
            link_recall.append(float(metrics["link_recall_at_10"]))
        if task.get("train_tasks_per_sec") is not None:
            throughput.append(float(task["train_tasks_per_sec"]))
        if task.get("train_context_examples_per_sec") is not None:
            context_throughput.append(float(task["train_context_examples_per_sec"]))
    return {
        "num_tasks": len(summaries),
        "num_ok": len(ok_tasks),
        "task_type_counts": task_type_counts,
        "classification": {
            "num_tasks": len(classification),
            "avg_auroc": finite_mean(classification),
            "avg_auroc_points": finite_mean(classification) * 100.0 if classification else None,
        },
        "regression": {
            "num_tasks": len(regression),
            "avg_mae": finite_mean(regression),
        },
        "multiclass": {
            "num_tasks": len(multiclass),
            "avg_mrr": finite_mean(multiclass),
        },
        "link_prediction": {
            "num_tasks": len(link_mrr),
            "avg_mrr_at_10": finite_mean(link_mrr),
            "avg_hit_at_10": finite_mean(link_hit),
            "avg_recall_at_10": finite_mean(link_recall),
        },
        "training": {
            "mean_tasks_per_sec": finite_mean(throughput),
            "mean_context_examples_per_sec": finite_mean(context_throughput),
        },
    }


def finite_mean(values: list[float]) -> float | None:
    finite = [value for value in values if value == value]
    return sum(finite) / len(finite) if finite else None


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        sys.exit(exc.returncode)
