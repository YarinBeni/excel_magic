"""P3 / claims C4, C7: prediction questions over RelBench databases answered by an agent, configs D, E, R, ER.

  python experiments/p3_relbench_capability.py --run DIR --tasks rel-f1/driver-dnf,... --configs D,E,R,ER
OPENAI_BASE_URL serves the agent LLM. Runs in the tabfm venv (relbench, sdm, tabfm_auto).
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd.llm import Chat  # noqa: E402
from vfd.relbench_env import QUESTIONS, RelEnv, run_agent  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--tasks", default=",".join(f"{d}/{t}" for d, t in QUESTIONS))
    ap.add_argument("--configs", default="D,E,R,ER")
    ap.add_argument("--repeats", type=int, default=2)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--deep-spec", default="kumo-tabular-l:n_estimators=4,device=cuda")
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    out = run / "capability_rows.jsonl"
    done = {(r["task"], r["config"], r["repeat"]) for r in map(json.loads, out.open())} if out.exists() else set()
    chat = Chat(a.model)
    for spec in a.tasks.split(","):
        ds, tn = spec.split("/")
        env = RelEnv(ds, tn).setup()
        for cfg in a.configs.split(","):
            for rep in range(a.repeats):
                if (spec, cfg, rep) in done:
                    continue
                env.reset()
                env.seed = rep
                try:
                    r = run_agent(env, chat, cfg, a.deep_spec)
                except Exception as e:
                    r = {"auroc": float("nan"), "submitted": False, "error": f"{type(e).__name__}: {str(e)[:200]}"}
                r.update(task=spec, config=cfg, repeat=rep, log=env.log[-30:])
                with out.open("a") as f:
                    f.write(json.dumps(r, default=str) + "\n")
                print(f"[cap] {spec} {cfg} rep{rep}: auroc {r.get('auroc')} submitted {r.get('submitted')} "
                      f"tools {r.get('tools')}", flush=True)
    df = pd.DataFrame([json.loads(x) for x in out.open()])
    piv = df.pivot_table(index="task", columns="config", values="auroc", aggfunc="mean")
    sub = df.pivot_table(index="task", columns="config", values="submitted", aggfunc="mean")
    md = ["# P3: prediction questions over RelBench databases, agent configs", "",
          "D = LLM + SQL; E = + Kumo Tabular-L fitted on LLM-built features; R = + Kumo Relational (graph layer probe); "
          "ER = both. Official test AUROC, mean over repeats (unsubmitted = missing).", "", piv.round(4).to_markdown(), "",
          "Share of episodes that submitted:", "", sub.round(2).to_markdown(), "",
          f"Mean over tasks: {piv.mean().round(4).to_dict()}"]
    (run / "results.md").write_text("\n".join(md) + "\n")
    json.dump({"mean_auroc": piv.mean().round(5).to_dict(), "per_task": piv.round(5).to_dict()},
              open(run / "metrics.json", "w"), indent=1, default=float)
    print("\n".join(md))


if __name__ == "__main__":
    main()
