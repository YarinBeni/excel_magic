"""P3 harness swap: the same InsightBench tables and the same three-model tools, driven by the pi coding agent
(the harness the company uses) instead of our tool loop. Tools reach pi as the `vfd` shell command, served by
`python -m vfd.toolserver --config D|E|F --port P` (one server per config, started by the job).

  python experiments/p3_insightbench_pi.py --run DIR --configs D,E,F --ports D=8765,E=8766,F=8767 --model M
Writes insights_pi<config>.jsonl in the same format as p3_insightbench.py, so its judge / report stages score it.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd.toolserver import HELP, HELP_DEEP, HELP_TEXT  # noqa: E402

PROMPT = """You are a senior data analyst working in this directory. The data is `data.csv` (also loaded as table
`data` in DuckDB behind the `vfd` command). Goal: {goal}
Role: {role}. Dataset: {desc}

{help}
Find concrete, data-backed insights about the goal: distributions, differences between groups, trends over time,
outliers, and what drives the outcome. Record each one with `vfd record_insight` (one sentence with the numbers, plus
the SQL that shows it). {verify}Record at most 8 insights, then stop."""


def post(port: int, ws: str, tool: str, args: dict | None = None) -> str:
    req = urllib.request.Request(f"http://127.0.0.1:{port}/", method="POST",
                                 data=json.dumps({"workspace": ws, "tool": tool, "args": args or {}}).encode())
    return urllib.request.urlopen(req, timeout=900).read().decode()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--bench", default=os.environ.get("INSIGHTBENCH", "artifacts/insight-bench"))
    ap.add_argument("--configs", default="D,E,F")
    ap.add_argument("--ports", default="D=8765,E=8766,F=8767")
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--timeout", type=int, default=900)
    a = ap.parse_args()
    run, root = Path(a.run), Path(a.bench)
    ports = dict(x.split("=") for x in a.ports.split(","))
    from p3_insightbench import flags  # same table list and order

    for cfg in a.configs.split(","):
        port = int(ports[cfg])
        out = run / f"insights_pi{cfg}.jsonl"
        done = {r["flag"] for r in map(json.loads, out.open())} if out.exists() else set()
        for fp in flags(root, a.limit):
            if fp.stem in done:
                continue
            d = json.load(open(fp))
            ws = run / f"ws_pi{cfg}" / fp.stem
            ws.mkdir(parents=True, exist_ok=True)
            shutil.copy2(root / d["dataset_csv_path"], ws / "data.csv")
            m = d["metadata"]
            help_txt = HELP.format(deep=HELP_DEEP if cfg in ("E", "F") else "", text=HELP_TEXT if cfg == "F" else "")
            verify = ("Insights are checked: every number you state must appear in the result of your evidence SQL; "
                      "rejected insights come back with the reason. " if cfg == "F" else "")
            (ws / "prompt.md").write_text(PROMPT.format(goal=m.get("goal", ""), role=m.get("role", "analyst"),
                                                        desc=str(m.get("dataset_description", ""))[:1500],
                                                        help=help_txt, verify=verify))
            env = {**os.environ, "VFD_PORT": str(port)}
            t0 = time.time()
            try:
                p = subprocess.run(["pi", "-a", "--provider", "vllm", "--model", a.model, "--mode", "json", "@prompt.md"],
                                   cwd=ws, env=env, capture_output=True, text=True, timeout=a.timeout)
                err = "" if p.returncode == 0 else f"exit {p.returncode}: {p.stderr[-300:]}"
            except subprocess.TimeoutExpired:
                err = "timeout"
            led = json.loads(post(port, str(ws), "ledger"))
            r = {"flag": fp.stem, "config": f"pi{cfg}", "insights": led["accepted"], "n_rejected": led["rejected"],
                 "seconds": round(time.time() - t0, 1), "error": err, "harness": "pi"}
            with out.open("a") as f:
                f.write(json.dumps(r) + "\n")
            print(f"[pi] {cfg} {fp.stem}: {len(r['insights'])} insights, {r['n_rejected']} rejected, "
                  f"{r['seconds']}s {err[:80]}", flush=True)


if __name__ == "__main__":
    main()
