"""Q3: LLM hypotheses tested on top of the frozen relational FM (vfd.hypo). Per task the FM is computed once; each
episode starts with no accepted hypothesis. Reports test AUROC of the system and of the FM alone (same rows).

  python experiments/p6_relbench_hypo.py --run DIR --tasks rel-f1/driver-dnf,... --repeats 2 --budget 10
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from vfd.hypo import HypoLoop  # noqa: E402
from vfd.llm import Chat  # noqa: E402
from vfd.relbench_env import QUESTIONS, RelEnv  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", required=True)
    ap.add_argument("--tasks", default=",".join(f"{d}/{t}" for d, t in QUESTIONS))
    ap.add_argument("--repeats", type=int, default=2)
    ap.add_argument("--budget", type=int, default=10)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--tag", default="")
    a = ap.parse_args()
    run = Path(a.run)
    run.mkdir(parents=True, exist_ok=True)
    out = run / "hypo_rows.jsonl"
    tag = a.tag or a.model.split("/")[-1]
    done = {(r["task"], r["repeat"], r["model"]) for r in map(json.loads, out.open())} if out.exists() else set()
    chat = Chat(a.model)
    for spec in a.tasks.split(","):
        ds, tn = spec.split("/")
        env = None
        for rep in range(a.repeats):
            if (spec, rep, tag) in done:
                continue
            if env is None:
                env = RelEnv(ds, tn).setup()
                env.relational_predict()
            try:
                r = HypoLoop(env, chat, budget=a.budget).run(QUESTIONS[(ds, tn)])
            except Exception as e:
                r = {"auroc": float("nan"), "error": f"{type(e).__name__}: {str(e)[:300]}"}
            r.update(task=spec, repeat=rep, model=tag)
            with out.open("a") as f:
                f.write(json.dumps(r, default=str) + "\n")
            print(f"[hypo] {spec} rep{rep}: system {r.get('auroc')} fm {r.get('auroc_fm')} accepted "
                  f"{r.get('accepted')} tested {r.get('tested')}", flush=True)
    report(run)


def report(run: Path) -> None:
    df = pd.DataFrame([json.loads(x) for x in (run / "hypo_rows.jsonl").open()])
    df["gain"] = df["auroc"] - df["auroc_fm"]
    g = df.groupby(["model", "task"]).agg(system=("auroc", "mean"), fm=("auroc_fm", "mean"), gain=("gain", "mean"),
                                          accepted=("accepted", lambda s: np.mean([len(x) for x in s])),
                                          tested=("tested", "mean")).reset_index()
    lines = ["# Q3: LLM hypotheses on top of the frozen relational FM (RelBench, official test AUROC)", "",
             g.round(4).to_markdown(index=False), ""]
    for m, gm in g.groupby("model"):
        lines.append(f"{m}: mean system {gm.system.mean():.4f} vs FM {gm.fm.mean():.4f}; gain {gm.gain.mean():+.4f}; "
                     f"better on {int((gm.gain > 0).sum())} / worse on {int((gm.gain < 0).sum())} of {len(gm)} tasks")
    (run / "results.md").write_text("\n".join(lines) + "\n")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
