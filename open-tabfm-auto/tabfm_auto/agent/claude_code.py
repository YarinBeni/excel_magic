"""Thin wrapper around the headless Claude Code CLI (`claude -p`) used as the coding-agent harness,
as in the paper's "Claude Code" configuration. Every stream event is appended to a JSONL log."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
from pathlib import Path
from typing import Any

DEFAULT_ALLOWED_TOOLS = (
    "Read,Edit,Write,MultiEdit,Glob,Grep,"
    "Bash(tabfm-eval:*),Bash(python:*),Bash(python3:*),Bash(cat:*),Bash(ls:*),Bash(head:*),Bash(tail:*),"
    "Bash(wc:*),Bash(cp:*),Bash(diff:*),Bash(echo:*)"
)


def claude_available() -> bool:
    return shutil.which("claude") is not None


def run_claude_code(prompt: str, cwd: Path, log_path: Path, model: str = "sonnet", max_turns: int = 80,
                    timeout_s: int = 3600, allowed_tools: str = DEFAULT_ALLOWED_TOOLS,
                    system_prompt: str | None = None, env_extra: dict[str, str] | None = None) -> dict[str, Any]:
    """Run one headless session. Returns {"rc", "elapsed_s", "result", "n_events", "usage"...}."""
    cmd = ["claude", "-p", prompt, "--model", model, "--max-turns", str(max_turns),
           "--output-format", "stream-json", "--verbose", "--allowedTools", allowed_tools,
           "--permission-mode", "acceptEdits"]
    if system_prompt:
        cmd += ["--append-system-prompt", system_prompt]
    env = dict(os.environ)
    env.update(env_extra or {})
    # the agent runs with our venv's PATH so that `tabfm-eval` resolves
    t0 = time.time()
    n_events, result, usage, cost = 0, None, None, None
    tool_counts: dict[str, int] = {}
    with open(log_path, "a", encoding="utf-8") as logf:
        try:
            p = subprocess.Popen(cmd, cwd=str(cwd), env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 text=True, bufsize=1)
            assert p.stdout is not None
            deadline = t0 + timeout_s
            for line in p.stdout:
                logf.write(line)
                logf.flush()
                n_events += 1
                try:
                    ev = json.loads(line)
                except json.JSONDecodeError:
                    continue
                if ev.get("type") == "assistant":
                    for blk in ev.get("message", {}).get("content", []):
                        if blk.get("type") == "tool_use":
                            tool_counts[blk.get("name", "?")] = tool_counts.get(blk.get("name", "?"), 0) + 1
                if ev.get("type") == "result":
                    result = ev.get("result")
                    usage = ev.get("usage")
                    cost = ev.get("total_cost_usd")
                if time.time() > deadline:
                    p.kill()
                    logf.write(json.dumps({"type": "harness", "event": "timeout", "timeout_s": timeout_s}) + "\n")
                    break
            rc = p.wait(timeout=60)
            err = p.stderr.read() if p.stderr else ""
            if err.strip():
                logf.write(json.dumps({"type": "harness", "event": "stderr", "text": err[-4000:]}) + "\n")
        except FileNotFoundError:
            return {"rc": 127, "error": "claude CLI not found", "elapsed_s": 0}
    return {"rc": rc, "elapsed_s": round(time.time() - t0, 1), "result": result, "n_events": n_events,
            "usage": usage, "total_cost_usd": cost, "tool_counts": tool_counts}
