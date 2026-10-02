"""TabFM-Auto search driver.

For one dataset:
  1. hold out a test split (never written into the agent workspace),
  2. build a workspace with train.parquet / task.json / metadata.json / pipeline.py (identity P0) / TASK.md,
  3. evaluate P0 with `tabfm-eval` (3-fold CV on the training split),
  4. let the coding agent (headless Claude Code) edit pipeline.py and call `tabfm-eval` under a budget,
  5. freeze the best-CV candidate, score it (and P0) once on the held-out split, write metrics.json.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

import numpy as np
from sklearn.model_selection import train_test_split

from ..data import TabularTask
from ..harness.evaluator import evaluate_holdout
from ..harness.metrics import primary_metric_name
from ..logging_utils import RunLogger, dumps
from ..pipeline import IDENTITY_PIPELINE
from . import prompts
from .claude_code import claude_available, run_claude_code


def make_split(task: TabularTask, test_size: float, seed: int):
    strat = task.y if task.task_type != "regression" else None
    idx = np.arange(len(task.X))
    tr, te = train_test_split(idx, test_size=test_size, random_state=seed, stratify=strat)
    return np.sort(tr), np.sort(te)


def build_workspace(run: RunLogger, task: TabularTask, tr_idx: np.ndarray, model_spec: str, budget_evals: int,
                    budget_minutes: int, max_rows: int, n_folds: int, seed: int, eval_timeout_s: int) -> Path:
    ws = run.run_dir / "workspace"
    ws.mkdir(exist_ok=True)
    df = task.X.iloc[tr_idx].reset_index(drop=True).copy()
    df[task.y.name] = task.y.iloc[tr_idx].to_numpy()
    df.to_parquet(ws / "train.parquet", index=False)
    task_json = {"dataset": task.name, "task_type": task.task_type, "target": task.y.name,
                 "metric": primary_metric_name(task.task_type), "model_spec": model_spec, "n_folds": n_folds,
                 "seed": seed, "max_rows": max_rows, "budget_evals": budget_evals, "budget_minutes": budget_minutes,
                 "eval_timeout_s": eval_timeout_s, "n_train": int(len(df)), "n_cols": int(task.X.shape[1])}
    (ws / "task.json").write_text(dumps(task_json, indent=2))
    (ws / "metadata.json").write_text(dumps(task.metadata, indent=2))
    (ws / "pipeline.py").write_text(IDENTITY_PIPELINE)
    (ws / "candidates").mkdir(exist_ok=True)
    return ws


def tabfm_eval(ws: Path) -> dict[str, Any]:
    """Call the harness CLI exactly like the agent does and return the appended record."""
    subprocess.run([sys.executable, "-m", "tabfm_auto.harness.cli", "--workspace", str(ws)], cwd=str(ws),
                   capture_output=True, text=True)
    recs = read_evals(ws)
    return recs[-1] if recs else {"status": "error", "error": "no eval record"}


def read_evals(ws: Path) -> list[dict[str, Any]]:
    p = ws / "evals.jsonl"
    if not p.exists():
        return []
    return [json.loads(l) for l in p.read_text().splitlines() if l.strip()]


def run_search(task: TabularTask, model_spec: str = "tabpfn", harness: str = "claude-code", llm_model: str = "sonnet",
               budget_evals: int = 12, budget_minutes: int = 30, max_turns: int = 80, test_size: float = 0.3,
               n_folds: int = 3, seed: int = 0, max_rows: int = 10000, eval_timeout_s: int = 900,
               baselines: tuple[str, ...] = ("hgb",), name: str | None = None, run_dir: Path | None = None,
               split: tuple[np.ndarray, np.ndarray] | None = None, llm_base_url: str | None = None,
               agent_cmd: str | None = None) -> dict[str, Any]:
    """Run one TabFM-Auto search. ``split`` = (train_idx, test_idx) overrides the random hold-out (e.g. an
    official benchmark split); the agent only ever sees ``train_idx`` rows."""
    cfg = {k: v for k, v in locals().items() if k not in ("task", "run_dir", "split")}
    cfg.update({"dataset": task.name, "model": model_spec, "official_split": split is not None, **task.summary()})
    name = name or f"search_{task.name}_{model_spec.split(':')[0]}_{harness}"
    with RunLogger(name, cfg, run_dir=run_dir) as run:
        tr_idx, te_idx = (np.asarray(split[0]), np.asarray(split[1])) if split is not None else make_split(task, test_size, seed)
        run.save_json("split.json", {"train_idx": tr_idx, "test_idx": te_idx})
        heldout = run.run_dir / "_heldout"
        heldout.mkdir(exist_ok=True)
        dte = task.X.iloc[te_idx].reset_index(drop=True).copy()
        dte[task.y.name] = task.y.iloc[te_idx].to_numpy()
        dte.to_parquet(heldout / "test.parquet", index=False)
        ws = build_workspace(run, task, tr_idx, model_spec, budget_evals, budget_minutes, max_rows, n_folds, seed,
                             eval_timeout_s)
        run.event("workspace_ready", path=str(ws), n_train=len(tr_idx), n_test=len(te_idx))

        # P0
        p0 = tabfm_eval(ws)
        run.event("p0_eval", **{k: v for k, v in p0.items() if k != "traceback"})
        run.info("P0 cv %s = %s", p0.get("metric"), p0.get("score"))
        (ws / "TASK.md").write_text(prompts.workspace_task_md(json.loads((ws / "task.json").read_text()),
                                                               task.metadata, p0))

        # agent
        agent_info: dict[str, Any] = {"harness": harness}
        if harness == "claude-code":
            if not claude_available():
                raise RuntimeError("claude CLI not found on PATH")
            prompt = (ws / "TASK.md").read_text()
            run.event("agent_start", model=llm_model, max_turns=max_turns, budget_minutes=budget_minutes)
            agent_info.update(run_claude_code(prompt, ws, run.run_dir / "agent_stream.jsonl", model=llm_model,
                                              max_turns=max_turns, timeout_s=budget_minutes * 60,
                                              system_prompt=prompts.SYSTEM_RULES))
            run.event("agent_end", **{k: v for k, v in agent_info.items() if k != "result"})
            run.save_text("agent_result.md", str(agent_info.get("result") or ""))
        elif harness == "openai":
            from .openai_compat import run_openai_agent

            prompt = (ws / "TASK.md").read_text()
            run.event("agent_start", model=llm_model, harness="openai", base_url=llm_base_url, max_turns=max_turns)
            agent_info.update(run_openai_agent(prompt, ws, run.run_dir / "agent_stream.jsonl", model=llm_model,
                                               base_url=llm_base_url, max_turns=max_turns, timeout_s=budget_minutes * 60,
                                               budget_evals=budget_evals, eval_timeout_s=eval_timeout_s,
                                               system_prompt=prompts.SYSTEM_RULES))
            run.event("agent_end", **{k: v for k, v in agent_info.items() if k != "result"})
            run.save_text("agent_result.md", str(agent_info.get("result") or ""))
        elif harness == "cli":
            from .cli_agent import run_cli_agent

            if not agent_cmd:
                raise ValueError("harness='cli' needs agent_cmd (a preset name like 'pi' or a command template)")
            prompt = (ws / "TASK.md").read_text()
            run.event("agent_start", harness="cli", agent_cmd=agent_cmd, model=llm_model, base_url=llm_base_url)
            agent_info.update(run_cli_agent(prompt, ws, run.run_dir / "agent_stream.log", agent_cmd, model=llm_model,
                                            base_url=llm_base_url, timeout_s=budget_minutes * 60))
            run.event("agent_end", **agent_info)
        elif harness == "heuristic":
            from .heuristic import run_heuristic_search

            run.event("agent_start", harness="heuristic", budget_evals=budget_evals, budget_minutes=budget_minutes)
            agent_info.update(run_heuristic_search(ws, task.task_type, budget_evals, budget_minutes, log=run,
                                                   eval_timeout_s=eval_timeout_s))
            run.event("agent_end", **{k: v for k, v in agent_info.items() if k != "trace"})
            run.save_json("heuristic_trace.json", agent_info.pop("trace", []))
        elif harness == "none":
            run.info("harness=none: only P0 is evaluated")
        else:
            raise ValueError(f"unknown harness {harness!r}; use claude-code | openai | cli | heuristic | none")

        # pick the best candidate by CV
        evals = read_evals(ws)
        ok = [r for r in evals if r.get("status") == "ok" and r.get("score") is not None]
        best = min(ok, key=lambda r: r["score"]) if ok else None
        metrics: dict[str, Any] = {"dataset": task.name, "task_type": task.task_type, "model": model_spec,
                                   "metric": primary_metric_name(task.task_type), "harness": harness,
                                   "llm_model": llm_model if harness in ("claude-code", "openai", "cli") else None,
                                   "n_evals": len(evals), "n_evals_ok": len(ok),
                                   "p0_cv": p0.get("score"), "best_cv": best["score"] if best else None,
                                   "best_candidate": best["candidate"] if best else None,
                                   "agent": {k: v for k, v in agent_info.items() if k != "result"}}
        # held-out scoring (fit on the full training split)
        Xtr, ytr = task.X.iloc[tr_idx].reset_index(drop=True), task.y.iloc[tr_idx].reset_index(drop=True)
        Xte, yte = task.X.iloc[te_idx].reset_index(drop=True), task.y.iloc[te_idx].reset_index(drop=True)
        p0_path = ws / "candidates" / "eval_001.py"
        if p0_path.exists():
            r = evaluate_holdout(p0_path, Xtr, ytr, Xte, yte, task.task_type, model_spec, seed, max_rows, run)
            metrics["p0_test"] = r.get("score")
            metrics["p0_test_secondary"] = r.get("secondary")
            run.event("holdout_p0", **{k: v for k, v in r.items() if k != "traceback"})
        if best:
            shutil.copy(ws / "candidates" / best["candidate"], run.run_dir / "best_pipeline.py")
            r = evaluate_holdout(run.run_dir / "best_pipeline.py", Xtr, ytr, Xte, yte, task.task_type, model_spec,
                                 seed, max_rows, run)
            metrics["best_test"] = r.get("score")
            metrics["best_test_secondary"] = r.get("secondary")
            metrics["best_test_info"] = r.get("info")
            run.event("holdout_best", **{k: v for k, v in r.items() if k != "traceback"})
            if metrics.get("p0_test") and metrics["best_test"] is not None:
                metrics["test_improvement_pct"] = round(100 * (metrics["p0_test"] - metrics["best_test"]) / metrics["p0_test"], 3)
        # classical baselines on the same split (same raw columns)
        for b in baselines:
            try:
                r = evaluate_holdout(p0_path, Xtr, ytr, Xte, yte, task.task_type, b, seed, max_rows, run)
                metrics[f"baseline_{b}_test"] = r.get("score")
            except Exception as e:
                metrics[f"baseline_{b}_test"] = f"error: {e}"
        metrics["cv_trajectory"] = [(r.get("eval_id"), r.get("score")) for r in evals]
        run.finish(metrics)
        return metrics
