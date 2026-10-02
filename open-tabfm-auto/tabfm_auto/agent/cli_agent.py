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
    # verified against upstream docs on 2026-10-02 (docs/LLM_BACKENDS.md has the provider setup for each)
    "pi": "pi -a --provider vllm --model {model} --mode json @{prompt_file}",
    "qwen-code": "qwen -p \"$(cat {prompt_file})\" --approval-mode yolo --output-format json --max-session-turns 60 --max-wall-time 50m -m {model}",
    "gemini-cli": "gemini -p \"$(cat {prompt_file})\" --approval-mode yolo --output-format json",
    "aider": "aider --message-file {prompt_file} --yes-always --no-git --no-auto-commits --no-show-model-warnings --no-check-update --no-analytics --model openai/{model} pipeline.py",
    "codex": "codex exec --dangerously-bypass-approvals-and-sandbox --skip-git-repo-check -C . -m {model} --json \"$(cat {prompt_file})\"",
    "openhands": "openhands --headless --json --override-with-envs -f {prompt_file}",
    "claude-code": "claude -p \"$(cat {prompt_file})\" --model {model} --permission-mode acceptEdits --allowedTools Read,Edit,Write,Bash --max-turns 60 --output-format json",
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
    if base_url:  # the env names the different CLIs read
        env.setdefault("OPENAI_BASE_URL", base_url)   # Qwen Code
        env.setdefault("OPENAI_API_BASE", base_url)   # aider
        env.setdefault("LLM_BASE_URL", base_url)      # OpenHands
        env.setdefault("OPENAI_API_KEY", env.get("OPENAI_API_KEY", "EMPTY"))
        env.setdefault("LLM_API_KEY", "EMPTY")
        if model:
            env.setdefault("OPENAI_MODEL", model)
            env.setdefault("LLM_MODEL", f"openai/{model}")
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
