"""Generic harness for any coding-agent CLI that can run headlessly in a directory (pi, Qwen Code, Gemini CLI,
aider, Codex CLI, OpenHands, ...).

The agent is started with a shell command template inside the workspace; it reads TASK.md, edits pipeline.py and
calls `tabfm-eval` like every other harness. Placeholders in the template:
  {prompt_file}  path of TASK.md          {prompt}  the brief as one shell-quoted argument
  {model}        --llm value              {base_url} --llm-base-url value (or "")
Budget is enforced by `tabfm-eval` (eval count) and by the wall-clock timeout here. Presets (PRESETS) are starting
points; CLIs change their flags, check `docs/LLM_BACKENDS.md`.
"""
from __future__ import annotations

import os
import shlex
import subprocess
import time
from pathlib import Path
from typing import Any

PRESETS: dict[str, str] = {
    # verified against upstream READMEs on 2026-10-02 where possible; see docs/LLM_BACKENDS.md
    "pi": "pi -p {prompt}",
    "qwen-code": "qwen -p {prompt} --yolo",
    "gemini-cli": "gemini -p {prompt} --yolo",
    "aider": "aider --yes-always --no-git --no-auto-commits --model openai/{model} --message {prompt} pipeline.py",
    "codex": "codex exec --full-auto --skip-git-repo-check -m {model} {prompt}",
    "claude-code": "claude -p {prompt} --model {model} --permission-mode acceptEdits --allowedTools Read,Edit,Write,Bash",
}


def run_cli_agent(prompt: str, ws: Path, log_path: Path, agent_cmd: str, model: str = "", base_url: str | None = None,
                  timeout_s: int = 3600, env_extra: dict[str, str] | None = None) -> dict[str, Any]:
    """Run ``agent_cmd`` (a preset name or a template) in ``ws``; stdout+stderr go to ``log_path``."""
    template = PRESETS.get(agent_cmd, agent_cmd)
    prompt_file = ws / "TASK.md"
    if not prompt_file.exists():
        prompt_file.write_text(prompt)
    cmd = template.format(prompt_file=shlex.quote(str(prompt_file)), prompt=shlex.quote(prompt),
                          model=shlex.quote(model) if model else "", base_url=shlex.quote(base_url) if base_url else "")
    env = dict(os.environ)
    if base_url:
        env.setdefault("OPENAI_BASE_URL", base_url)
        env.setdefault("OPENAI_API_BASE", base_url)
        env.setdefault("OPENAI_API_KEY", env.get("OPENAI_API_KEY", "EMPTY"))
    env.update(env_extra or {})
    t0 = time.time()
    with open(log_path, "a", encoding="utf-8") as logf:
        logf.write(f"# cmd: {cmd}\n# cwd: {ws}\n")
        logf.flush()
        try:
            p = subprocess.run(cmd, shell=True, cwd=str(ws), env=env, stdout=logf, stderr=subprocess.STDOUT,
                               timeout=timeout_s, stdin=subprocess.DEVNULL, check=False)
            rc = p.returncode
        except subprocess.TimeoutExpired:
            rc = 124
            logf.write(f"\n# harness: timeout after {timeout_s}s\n")
    return {"rc": rc, "elapsed_s": round(time.time() - t0, 1), "agent_cmd": cmd, "model": model, "harness": "cli"}
