"""Pipeline-search agent on any OpenAI-compatible chat endpoint (vLLM / SGLang / Ollama / OpenAI / OpenRouter).

Use it to run the search with an open-weights coding model on your own cluster, e.g.

    vllm serve Qwen/Qwen3-Coder-30B-A3B-Instruct --enable-auto-tool-choice --tool-call-parser hermes --port 8000
    tabfm-auto search --dataset synth_physics --harness openai --llm Qwen/Qwen3-Coder-30B-A3B-Instruct \
        --llm-base-url http://localhost:8000/v1

The loop is deliberately small: the model gets the task brief, four tools (read_file, describe_data,
write_pipeline, run_eval), and loops until it says it is done, the eval budget is spent, or max_turns is hit.
Every request/response is appended to a JSONL log. Requires the ``openai`` package.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

TOOLS = [
    {"type": "function", "function": {"name": "read_file", "description": "Read a file in the workspace (pipeline.py, TASK.md, NOTES.md, evals.jsonl, candidates/...).",
                                      "parameters": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}}},
    {"type": "function", "function": {"name": "describe_data", "description": "Summary statistics of train.parquet: dtypes, describe(), head, target distribution.",
                                      "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "write_pipeline", "description": "Overwrite pipeline.py with the full new source (must keep the four hooks + MODEL_KWARGS).",
                                      "parameters": {"type": "object", "properties": {"source": {"type": "string"}}, "required": ["source"]}}},
    {"type": "function", "function": {"name": "run_eval", "description": "Score the current pipeline.py with 3-fold CV (tabfm-eval). Returns score, metric, errors, budget left.",
                                      "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "finish", "description": "Stop searching. pipeline.py must equal your best candidate; give a 3-6 line summary in notes.",
                                      "parameters": {"type": "object", "properties": {"notes": {"type": "string"}}, "required": ["notes"]}}},
]


def _describe(ws: Path) -> str:
    import pandas as pd

    task = json.loads((ws / "task.json").read_text())
    df = pd.read_parquet(ws / "train.parquet")
    out = [f"shape={df.shape} target={task['target']} task_type={task['task_type']}", "dtypes:",
           df.dtypes.astype(str).to_string(), "describe:", df.describe(include="all").T.head(60).to_string(),
           "head:", df.head(5).to_string()]
    if task["task_type"] != "regression":
        out += ["target distribution:", df[task["target"]].value_counts(normalize=True).to_string()]
    return "\n".join(out)[:12000]


def _run_eval(ws: Path, timeout_s: int) -> str:
    p = subprocess.run([sys.executable, "-m", "tabfm_auto.harness.cli", "--workspace", str(ws)], cwd=str(ws),
                       capture_output=True, text=True, timeout=timeout_s + 60, check=False)
    return (p.stdout + ("\n" + p.stderr[-1500:] if p.returncode not in (0, 1) else ""))[-6000:]


def _clean_tool_name(name: str) -> str:
    """gpt-oss (harmony) through vLLM's parser leaks channel markers into the tool name
    (``run_eval<|channel|>commentary``, ``run_eval..commentary``); keep the identifier only."""
    name = (name or "").strip()
    for sep in ("<|", "..", "<", " ", "\n"):
        name = name.split(sep)[0]
    return name.strip()


_BINARY_SUFFIXES = {".parquet", ".npy", ".npz", ".pkl", ".pickle", ".feather", ".arrow", ".zip", ".gz", ".pt", ".bin"}


def execute_tool(name: str, args: dict[str, Any], ws: Path, eval_timeout_s: int) -> str:
    """Run one tool call. Every failure comes back to the model as an ``error: ...`` string; a tool bug or a bad
    argument must never abort the whole search (gpt-oss read a parquet file and the UnicodeDecodeError killed the run)."""
    try:
        return _execute_tool(name, args, ws, eval_timeout_s)
    except Exception as e:
        return f"error: {type(e).__name__}: {str(e)[:500]}"


def _execute_tool(name: str, args: dict[str, Any], ws: Path, eval_timeout_s: int) -> str:
    if name == "read_file":
        p = (ws / str(args.get("path", ""))).resolve()
        if ws.resolve() not in p.parents and p != ws.resolve():
            return "error: path outside the workspace"
        if not p.exists():
            return f"error: {args.get('path')} not found"
        if p.suffix.lower() in _BINARY_SUFFIXES:
            return f"error: {p.name} is a binary data file; use describe_data for the schema and summary statistics"
        return p.read_text(errors="replace")[-20000:]
    if name == "describe_data":
        return _describe(ws)
    if name == "write_pipeline":
        src = str(args.get("source", ""))
        for hook in ("def preprocess", "def engineer", "def sample", "def postprocess"):
            if hook not in src:
                return f"error: pipeline.py must define {hook}(...)"
        (ws / "pipeline.py").write_text(src)
        return f"pipeline.py written ({len(src)} chars). Call run_eval to score it."
    if name == "run_eval":
        # a run_eval without a change in between burns budget for an identical score (Qwen3-Coder looped 38 times)
        src = (ws / "pipeline.py").read_bytes() if (ws / "pipeline.py").exists() else b""
        digest = hashlib.sha1(src).hexdigest()
        marker = ws / ".last_eval_sha1"
        if marker.exists() and marker.read_text() == digest:
            return "error: pipeline.py is unchanged since the last run_eval (same score would come back); call write_pipeline with a different pipeline first, or finish"
        out = _run_eval(ws, eval_timeout_s)
        marker.write_text(digest)
        return out
    if name == "finish":
        (ws / "NOTES.md").write_text(str(args.get("notes", "")))
        return "done"
    return f"error: unknown tool {name}"


def run_openai_agent(prompt: str, ws: Path, log_path: Path, model: str, base_url: str | None = None,
                     api_key: str | None = None, max_turns: int = 80, timeout_s: int = 3600, budget_evals: int | None = None,
                     eval_timeout_s: int = 900, system_prompt: str | None = None, temperature: float = 0.2,
                     client: Any = None, rich: bool = False) -> dict[str, Any]:
    """Returns {"rc", "elapsed_s", "result", "n_turns", "tool_counts", "usage"}. ``client`` may be injected (tests).
    ``rich`` prepends what a CLI coding agent sees anyway (file listing, pipeline.py, the data summary) to the first
    message, to test whether the harness gap is affordances rather than agency."""
    if client is None:
        from openai import OpenAI

        client = OpenAI(base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
                        api_key=api_key or os.environ.get("OPENAI_API_KEY", "EMPTY"))
    messages: list[dict[str, Any]] = []
    if system_prompt:
        messages.append({"role": "system", "content": system_prompt})
    if rich:
        try:
            listing = "\n".join(sorted(p.name for p in ws.iterdir()))
            pipe = (ws / "pipeline.py").read_text()[:6000] if (ws / "pipeline.py").exists() else ""
            prompt = (f"{prompt}\n\n## Workspace files\n{listing}\n\n## Current pipeline.py\n```python\n{pipe}\n```\n\n"
                      f"## Data summary (train.parquet)\n{_describe(ws)[:12000]}\n\nStart by calling write_pipeline with "
                      "your first candidate, then run_eval; iterate on the score.")
        except Exception as e:  # never let the preamble break the run
            prompt = f"{prompt}\n\n(rich context unavailable: {type(e).__name__}: {e})"
    messages.append({"role": "user", "content": prompt})
    t0 = time.time()
    tool_counts: dict[str, int] = {}
    usage_in = usage_out = 0
    result = None
    rc = 0
    with open(log_path, "a", encoding="utf-8") as logf:
        for turn in range(max_turns):
            if time.time() - t0 > timeout_s:
                rc = 124
                break
            try:
                resp = client.chat.completions.create(model=model, messages=messages, tools=TOOLS, temperature=temperature)
            except Exception as e:
                logf.write(json.dumps({"turn": turn, "error": repr(e)}) + "\n")
                rc = 1
                break
            if not getattr(resp, "choices", None):  # vLLM can answer with choices=None on a parser/length failure
                logf.write(json.dumps({"turn": turn, "error": f"empty response: {str(resp)[:500]}"}) + "\n")
                rc = 1
                break
            msg = resp.choices[0].message
            u = getattr(resp, "usage", None)
            if u is not None:
                usage_in += getattr(u, "prompt_tokens", 0) or 0
                usage_out += getattr(u, "completion_tokens", 0) or 0
            entry: dict[str, Any] = {"turn": turn, "content": msg.content, "tool_calls": []}
            messages.append({"role": "assistant", "content": msg.content or "",
                             "tool_calls": [tc.model_dump() if hasattr(tc, "model_dump") else tc for tc in (msg.tool_calls or [])]})
            if not msg.tool_calls:
                result = msg.content
                logf.write(json.dumps(entry) + "\n")
                break
            finished = False
            for tc in msg.tool_calls:
                name = _clean_tool_name(tc.function.name)
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except json.JSONDecodeError:
                    args = {}
                tool_counts[name] = tool_counts.get(name, 0) + 1
                out = execute_tool(name, args, ws, eval_timeout_s)
                entry["tool_calls"].append({"name": name, "args": {k: (v if k != "source" else f"<{len(v)} chars>") for k, v in args.items()},
                                            "result": out[-2000:]})
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": out})
                if name == "finish":
                    finished, result = True, args.get("notes")
                if name == "run_eval" and "BUDGET EXHAUSTED" in out:
                    messages.append({"role": "user", "content": "Budget exhausted. Restore your best candidate to pipeline.py "
                                                                "(copy from candidates/) and call finish."})
            logf.write(json.dumps(entry) + "\n")
            if finished:
                break
    return {"rc": rc, "elapsed_s": round(time.time() - t0, 1), "result": result, "n_turns": len([m for m in messages if m["role"] == "assistant"]),
            "tool_counts": tool_counts, "usage": {"input_tokens": usage_in, "output_tokens": usage_out}, "model": model}
