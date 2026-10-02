"""Run a candidate evaluation in a child process with network disabled and a wall-clock limit.

We cannot rely on bubblewrap in every environment, so isolation is best-effort: the child gets an
environment with every proxy variable removed and ``no_proxy='*'``, offline flags for model hubs,
a CPU-thread cap, and a timeout. Held-out test rows are never placed inside the workspace.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

PROXY_VARS = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy", "ALL_PROXY", "all_proxy")


def sandbox_env(threads: int = 4) -> dict[str, str]:
    env = {k: v for k, v in os.environ.items() if k not in PROXY_VARS}
    env.update({
        "no_proxy": "*", "NO_PROXY": "*",
        "HF_HUB_OFFLINE": "1", "TRANSFORMERS_OFFLINE": "1",
        "OMP_NUM_THREADS": str(threads), "MKL_NUM_THREADS": str(threads),
        "TOKENIZERS_PARALLELISM": "false", "PYTHONUNBUFFERED": "1",
    })
    return env


def run_isolated(args: list[str], cwd: Path, timeout_s: int, threads: int = 4) -> dict[str, Any]:
    """Run ``python -m ... args`` and return the JSON object printed on the child's last stdout line."""
    cmd = [sys.executable, *args]
    try:
        p = subprocess.run(cmd, cwd=str(cwd), env=sandbox_env(threads), capture_output=True, text=True,
                           timeout=timeout_s)
    except subprocess.TimeoutExpired as e:
        return {"status": "error", "score": None, "error": f"timeout after {timeout_s}s",
                "stderr": (e.stderr or "")[-2000:] if isinstance(e.stderr, str) else ""}
    lines = [ln for ln in p.stdout.strip().splitlines() if ln.strip()]
    if p.returncode != 0 or not lines:
        return {"status": "error", "score": None, "error": f"child exited {p.returncode}",
                "stderr": p.stderr[-3000:], "stdout": p.stdout[-1000:]}
    try:
        out = json.loads(lines[-1])
    except json.JSONDecodeError:
        return {"status": "error", "score": None, "error": "child printed no JSON", "stdout": p.stdout[-2000:],
                "stderr": p.stderr[-2000:]}
    if p.stderr.strip():
        out["stderr_tail"] = p.stderr[-1500:]
    return out
