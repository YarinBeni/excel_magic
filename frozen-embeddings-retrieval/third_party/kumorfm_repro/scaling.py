from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Run small KumoRFM synthetic scaling sweeps.")
    p.add_argument("--output-dir", type=Path, default=Path("runs/scaling"))
    p.add_argument("--context-sizes", type=int, nargs="+", default=[8, 16, 32])
    p.add_argument("--steps", type=int, default=20)
    p.add_argument("--batch-size", type=int, default=8)
    p.add_argument("--d-model", type=int, default=64)
    p.add_argument("--layers", type=int, default=1)
    p.add_argument("--nproc-per-node", type=int, default=2)
    p.add_argument("--task-type", choices=["binary", "regression", "multiclass", "mixed"], default="mixed")
    p.add_argument("--schema-dropout", type=float, default=0.25)
    p.add_argument("--amp", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    summaries = []
    for context_size in args.context_sizes:
        run_dir = args.output_dir / f"context_{context_size}"
        cmd = [
            "torchrun",
            "--standalone",
            f"--nproc_per_node={args.nproc_per_node}",
            "-m",
            "kumorfm_repro.train",
            "--steps",
            str(args.steps),
            "--batch-size",
            str(args.batch_size),
            "--context-size",
            str(context_size),
            "--rows-per-child",
            "4",
            "--rows-per-aux",
            "3",
            "--schema-dropout",
            str(args.schema_dropout),
            "--task-type",
            args.task_type,
            "--d-model",
            str(args.d_model),
            "--heads",
            "4",
            "--layers",
            str(args.layers),
            "--val-batches",
            "8",
            "--eval-batches",
            "4",
            "--output-dir",
            str(run_dir),
            "--save-checkpoint",
        ]
        if args.amp:
            cmd.append("--amp")
        print(" ".join(cmd), flush=True)
        if not args.dry_run:
            subprocess.run(cmd, check=True)
            manifest = json.loads((run_dir / "manifest.json").read_text())
            summaries.append(
                {
                    "context_size": context_size,
                    "validation_score": manifest.get("validation_score"),
                    "train_tasks_per_sec": manifest.get("train_tasks_per_sec"),
                    "train_context_examples_per_sec": manifest.get("train_context_examples_per_sec"),
                    "effective_batch_size": manifest.get("effective_batch_size"),
                    "world_size": manifest.get("world_size"),
                    "run_dir": str(run_dir),
                }
            )
    if not args.dry_run:
        (args.output_dir / "scaling_summary.json").write_text(json.dumps(summaries, indent=2, sort_keys=True) + "\n")
        print(json.dumps(summaries, indent=2, sort_keys=True))
    return None


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        sys.exit(exc.returncode)
