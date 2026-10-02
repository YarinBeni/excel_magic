"""Synthetic pre-training for the N-table (database-native) RFM."""
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
from torch import nn
from torch.nn.parallel import DistributedDataParallel
from tqdm import trange
from sklearn.metrics import roc_auc_score

from .multitable import MultiTableRFM, MultiTableTaskGenerator, multitable_task_loss
from .train import build_scheduler, json_safe, finite_gradients, setup_distributed, save_checkpoint


def parse_args():
    p = argparse.ArgumentParser(description="Pre-train the N-table database-native RFM on synthetic tasks.")
    p.add_argument("--steps", type=int, default=2000)
    p.add_argument("--batch-size", type=int, default=12)
    p.add_argument("--context-size", type=int, default=24)
    p.add_argument("--num-event-tables", type=int, default=5)
    p.add_argument("--rows-per-event", type=int, default=10)
    p.add_argument("--root-cols", type=int, default=8)
    p.add_argument("--event-cols", type=int, default=8)
    p.add_argument("--lag-steps", type=int, default=4)
    p.add_argument("--d-model", type=int, default=256)
    p.add_argument("--heads", type=int, default=8)
    p.add_argument("--table-layers", type=int, default=2)
    p.add_argument("--graph-layers", type=int, default=3)
    p.add_argument("--max-classes", type=int, default=8)
    p.add_argument("--task-type", default="mixed")
    p.add_argument("--missing-prob", type=float, default=0.1)
    p.add_argument("--link-dropout", type=float, default=0.15)
    p.add_argument("--table-dropout", type=float, default=0.3)
    p.add_argument("--binary-quantile-min", type=float, default=0.5)
    p.add_argument("--binary-quantile-max", type=float, default=0.5)
    p.add_argument("--root-shortcut-dropout", type=float, default=0.0)
    p.add_argument("--history-label-frac", type=float, default=0.0)
    p.add_argument("--typed-task-conditioning", action="store_true")
    p.add_argument("--lr", type=float, default=2e-4)
    p.add_argument("--weight-decay", type=float, default=1e-2)
    p.add_argument("--warmup-steps", type=int, default=100)
    p.add_argument("--min-lr-ratio", type=float, default=0.02)
    p.add_argument("--grad-accum-steps", type=int, default=1)
    p.add_argument("--amp", action="store_true")
    p.add_argument("--val-batches", type=int, default=32)
    p.add_argument("--seed", type=int, default=0)
    p.add_argument("--output-dir", type=Path, default=Path("runs/mt_smoke"))
    p.add_argument("--save-checkpoint", action="store_true")
    p.add_argument("--checkpoint-every", type=int, default=0)
    p.add_argument("--log-every", type=int, default=200)
    return p.parse_args()


@torch.no_grad()
def validate(model, gen, device, batches, batch_size):
    model.eval()
    outs, ys, tts = [], [], []
    for _ in range(batches):
        b = gen.batch(batch_size).to(device)
        o = model(b).float().cpu().numpy()
        outs.append(np.nan_to_num(o, nan=0.0, posinf=1e4, neginf=-1e4))
        ys.append(b.y.cpu().numpy()); tts.append(b.task_type.cpu().numpy())
    model.train()
    out = np.concatenate(outs); y = np.concatenate(ys); tt = np.concatenate(tts)
    binary = tt == 0
    if binary.sum() > 0 and len(np.unique(y[binary])) > 1:
        return float(roc_auc_score(y[binary], 1 / (1 + np.exp(-out[binary, 0]))))
    return float("nan")


def main():
    args = parse_args()
    distributed, rank, world, local_rank = setup_distributed()
    device = torch.device(f"cuda:{local_rank}" if torch.cuda.is_available() else "cpu")
    torch.manual_seed(args.seed + rank); np.random.seed(args.seed + rank)

    def make_gen(seed):
        return MultiTableTaskGenerator(
            context_size=args.context_size, root_cols=args.root_cols, event_cols=args.event_cols,
            num_event_tables=args.num_event_tables, rows_per_event=args.rows_per_event,
            lag_steps=args.lag_steps, missing_prob=args.missing_prob, link_dropout=args.link_dropout,
            table_dropout=args.table_dropout, task_type=args.task_type, max_classes=args.max_classes,
            binary_quantile_min=args.binary_quantile_min, binary_quantile_max=args.binary_quantile_max,
            root_shortcut_dropout=args.root_shortcut_dropout, history_label_frac=args.history_label_frac, seed=seed,
        )

    train_gen = make_gen(args.seed + 10_000 * rank)
    val_gen = make_gen(args.seed + 999)
    model = MultiTableRFM(
        root_cols=args.root_cols, event_cols=args.event_cols, num_event_tables=args.num_event_tables,
        lag_steps=args.lag_steps, max_classes=args.max_classes, d_model=args.d_model, heads=args.heads,
        table_layers=args.table_layers, graph_layers=args.graph_layers,
        typed_task_conditioning=args.typed_task_conditioning,
    ).to(device)
    if distributed:
        model = DistributedDataParallel(model, device_ids=[local_rank])
    raw = model.module if distributed else model
    opt = torch.optim.AdamW(model.parameters(), lr=args.lr, weight_decay=args.weight_decay)
    sched = build_scheduler(opt, args.steps, args.warmup_steps, args.min_lr_ratio)
    scaler = torch.amp.GradScaler("cuda", enabled=args.amp and device.type == "cuda")

    model.train()
    started = time(); last_loss = float("nan"); skipped = 0; completed = 0
    it = trange(args.steps, disable=rank != 0)
    for step in it:
        opt.zero_grad(set_to_none=True); accum = 0.0; finite = True
        for ai in range(args.grad_accum_steps):
            b = train_gen.batch(args.batch_size).to(device)
            sc = model.no_sync() if (distributed and ai < args.grad_accum_steps - 1) else nullcontext()
            with sc:
                with (torch.amp.autocast("cuda", enabled=args.amp) if device.type == "cuda" else nullcontext()):
                    out = model(b)
                    loss = multitable_task_loss(out, b) / args.grad_accum_steps
                fl = torch.isfinite(loss).float()
                if distributed: dist.all_reduce(fl, op=dist.ReduceOp.MIN)
                if fl.item() < 1.0: finite = False; break
                scaler.scale(loss).backward(); accum += float(loss.detach().cpu()) * args.grad_accum_steps
        if not finite:
            skipped += 1; opt.zero_grad(set_to_none=True)
            if rank == 0: it.set_postfix(loss="nonfinite-skip")
            continue
        scaler.unscale_(opt)
        fg = finite_gradients(raw)
        if distributed: dist.all_reduce(fg, op=dist.ReduceOp.MIN)
        if fg.item() < 1.0:
            skipped += 1; opt.zero_grad(set_to_none=True); scaler.update()
            if rank == 0: it.set_postfix(loss="nonfinite-grad-skip")
            continue
        nn.utils.clip_grad_norm_(model.parameters(), 1.0)
        old = scaler.get_scale(); scaler.step(opt); scaler.update()
        if not scaler.is_enabled() or scaler.get_scale() >= old: sched.step()
        completed += 1; last_loss = accum / max(1, args.grad_accum_steps)
        if rank == 0 and (step + 1) % args.log_every == 0:
            it.set_postfix(loss=f"{last_loss:.4f}", lr=f"{sched.get_last_lr()[0]:.2e}")
        if rank == 0 and args.checkpoint_every and (step + 1) % args.checkpoint_every == 0:
            _save(args, raw, opt, sched, scaler, step + 1, train_gen)

    val = validate(raw, val_gen, device, args.val_batches, args.batch_size)
    if distributed:
        m = torch.tensor([val if math.isfinite(val) else 0.0], device=device); dist.all_reduce(m, op=dist.ReduceOp.AVG); val = float(m.item())
    if rank == 0:
        print(f"validation_auroc={val:.4f} world={world} completed={completed}/{args.steps} skipped={skipped}")
        args.output_dir.mkdir(parents=True, exist_ok=True)
        manifest = {"validation_auroc": val, "optimizer_steps": completed, "global_step": args.steps,
                    "skipped_nonfinite_steps": skipped, "last_train_loss": last_loss, "world_size": world,
                    "train_tasks_per_sec": completed * args.batch_size * world * args.grad_accum_steps / max(1e-9, time() - started),
                    "args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()}}
        (args.output_dir / "manifest.json").write_text(json.dumps(json_safe(manifest), allow_nan=False, indent=2, sort_keys=True) + "\n")
        if args.save_checkpoint: _save(args, raw, opt, sched, scaler, args.steps, train_gen)
    if distributed: dist.destroy_process_group()


def _save(args, raw, opt, sched, scaler, step, gen):
    args.output_dir.mkdir(parents=True, exist_ok=True)
    torch.save({"model": raw.state_dict(), "manifest": {"args": {k: str(v) if isinstance(v, Path) else v for k, v in vars(args).items()}, "global_step": step},
                "global_step": step}, args.output_dir / (f"checkpoint_step_{step}.pt" if step != args.steps else "checkpoint.pt"))


if __name__ == "__main__":
    main()
