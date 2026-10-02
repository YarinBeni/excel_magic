"""Run directories and structured logging.

Every experiment creates ``runs/<UTC timestamp>_<name>/`` containing:
- ``config.json``   the full configuration used
- ``events.jsonl``  one JSON object per logged event (append-only)
- ``run.log``       human-readable log (python logging, DEBUG level)
- ``metrics.json``  final metrics, written by ``RunLogger.finish``
"""
from __future__ import annotations

import json
import logging
import os
import platform
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


def _default_runs_root() -> Path:
    """``$TABFM_RUNS_ROOT`` if set, else ``./runs`` in the current working directory."""
    return Path(os.environ.get("TABFM_RUNS_ROOT", Path.cwd() / "runs"))


RUNS_ROOT = _default_runs_root()


def _json_default(o: Any):
    try:
        import numpy as np

        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        if isinstance(o, np.ndarray):
            return o.tolist()
    except ImportError:  # pragma: no cover
        pass
    if isinstance(o, Path):
        return str(o)
    if isinstance(o, set):
        return sorted(o)
    return str(o)


def dumps(obj: Any, **kw) -> str:
    return json.dumps(obj, default=_json_default, **kw)


def git_sha(repo: Path | None = None) -> str | None:
    repo = repo or Path(__file__).resolve().parents[1]
    try:
        return subprocess.check_output(
            ["git", "-C", str(repo), "rev-parse", "--short", "HEAD"], text=True, stderr=subprocess.DEVNULL
        ).strip()
    except Exception:
        return None


def env_fingerprint() -> dict[str, Any]:
    info: dict[str, Any] = {
        "python": sys.version.split()[0],
        "platform": platform.platform(),
        "cpu_count": os.cpu_count(),
        "git_sha": git_sha(),
    }
    for mod in ("numpy", "pandas", "sklearn", "torch", "tabpfn", "tabicl", "torch_geometric", "lightgbm"):
        try:
            m = __import__(mod)
            info[mod] = getattr(m, "__version__", "?")
        except Exception:
            info[mod] = None
    try:
        import torch

        info["torch_cuda"] = torch.cuda.is_available()
        info["torch_threads"] = torch.get_num_threads()
    except Exception:
        pass
    return info


class RunLogger:
    """Context manager owning one run directory."""

    def __init__(self, name: str, config: dict[str, Any] | None = None, root: Path | None = None,
                 run_dir: Path | None = None):
        root = Path(root) if root else _default_runs_root()
        if run_dir is None:
            ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
            run_dir = root / f"{ts}_{name}"
        self.name = name
        self.run_dir = Path(run_dir)
        self.run_dir.mkdir(parents=True, exist_ok=True)
        self.config = dict(config or {})
        self.t0 = time.time()
        self._events = open(self.run_dir / "events.jsonl", "a", encoding="utf-8")
        self.log = logging.getLogger(f"tabfm.{name}")
        self.log.setLevel(logging.DEBUG)
        self.log.propagate = False
        fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
        fh = logging.FileHandler(self.run_dir / "run.log", encoding="utf-8")
        fh.setFormatter(fmt)
        fh.setLevel(logging.DEBUG)
        sh = logging.StreamHandler(sys.stderr)
        sh.setFormatter(fmt)
        sh.setLevel(logging.INFO)
        self.log.handlers = [fh, sh]
        (self.run_dir / "config.json").write_text(
            dumps({"name": name, "config": self.config, "env": env_fingerprint(),
                   "started_utc": datetime.now(timezone.utc).isoformat()}, indent=2)
        )
        self.event("run_start", name=name)

    # -- events ---------------------------------------------------------------------------------
    def event(self, kind: str, **payload: Any) -> None:
        rec = {"t": round(time.time() - self.t0, 3), "utc": datetime.now(timezone.utc).isoformat(),
               "kind": kind, **payload}
        self._events.write(dumps(rec) + "\n")
        self._events.flush()
        short = {k: v for k, v in payload.items() if not isinstance(v, (dict, list)) or len(str(v)) < 200}
        self.log.debug("%s %s", kind, dumps(short))

    def info(self, msg: str, *args: Any) -> None:
        self.log.info(msg, *args)

    def warning(self, msg: str, *args: Any) -> None:
        self.log.warning(msg, *args)

    # -- artifacts ------------------------------------------------------------------------------
    def save_json(self, name: str, obj: Any) -> Path:
        p = self.run_dir / name
        p.write_text(dumps(obj, indent=2))
        return p

    def save_text(self, name: str, text: str) -> Path:
        p = self.run_dir / name
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text)
        return p

    def finish(self, metrics: dict[str, Any], status: str = "ok") -> Path:
        metrics = dict(metrics)
        metrics.setdefault("elapsed_s", round(time.time() - self.t0, 2))
        metrics["status"] = status
        p = self.save_json("metrics.json", metrics)
        self.event("run_end", status=status, elapsed_s=metrics["elapsed_s"])
        self.info("run finished (%s) -> %s", status, p)
        self._events.close()
        return p

    def __enter__(self) -> RunLogger:
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if exc is not None:
            self.log.exception("run failed: %s", exc)
            try:
                self.finish({"error": repr(exc)}, status="error")
            except Exception:
                pass


# -- summary CLI --------------------------------------------------------------------------------
def list_runs(root: Path | None = None) -> list[dict[str, Any]]:
    root = Path(root) if root else _default_runs_root()
    rows = []
    for d in sorted(root.glob("*_*")):
        if not d.is_dir():
            continue
        row: dict[str, Any] = {"run": d.name}
        cfg = d / "config.json"
        met = d / "metrics.json"
        if cfg.exists():
            try:
                row["config"] = json.loads(cfg.read_text()).get("config", {})
            except Exception:
                row["config"] = {}
        if met.exists():
            try:
                row["metrics"] = json.loads(met.read_text())
            except Exception:
                row["metrics"] = {}
        else:
            row["metrics"] = {"status": "incomplete"}
        rows.append(row)
    return rows


def summarize_cli(argv: list[str] | None = None) -> None:
    import argparse

    ap = argparse.ArgumentParser(description="Summarize run directories as a markdown table.")
    ap.add_argument("--root", default=None)
    ap.add_argument("--keys", default="status,dataset,model,metric,p0_cv,best_cv,p0_test,best_test,n_evals,elapsed_s",
                    help="comma-separated metric/config keys to show")
    a = ap.parse_args(argv)
    keys = [k for k in a.keys.split(",") if k]
    rows = list_runs(a.root)
    print("| run | " + " | ".join(keys) + " |")
    print("|---|" + "---|" * len(keys))
    for r in rows:
        vals = []
        for k in keys:
            v = r["metrics"].get(k, r.get("config", {}).get(k, ""))
            if isinstance(v, float):
                v = f"{v:.4g}"
            vals.append(str(v))
        print(f"| {r['run']} | " + " | ".join(vals) + " |")


if __name__ == "__main__":  # pragma: no cover
    summarize_cli()
