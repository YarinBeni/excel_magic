"""P3 / claim C3: InsightBench (ServiceNow, ICLR 2025; 100 business tables with planted insights).

  analyze  (tabfm venv: DEEP tool + GLiClass; generator LLM on OPENAI_BASE_URL)  configs D, E, F per table
  judge    (judge LLM on OPENAI_BASE_URL, another family)  open G-Eval-style match per gold insight + ROUGE-1
  report
D = LLM + SQL; E = + DEEP tools; F = + GLiClass text labels + verified insight ledger.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def flags(root: Path, limit: int = 0) -> list[Path]:
    fs = sorted((root / "data" / "notebooks").glob("flag-*.json"), key=lambda p: int(re.findall(r"\d+", p.stem)[0]))
    return fs[:limit] if limit else fs


def rouge1(pred: list[str], gold: list[str]) -> float:
    def toks(s):
        return re.findall(r"[a-z0-9]+", s.lower())

    def f1(a, b):
        A, B = toks(a), toks(b)
        if not A or not B:
            return 0.0
        common = sum(min(A.count(w), B.count(w)) for w in set(A))
        if not common:
            return 0.0
        p, r = common / len(A), common / len(B)
        return 2 * p * r / (p + r)

    return float(np.mean([max((f1(p, g) for p in pred), default=0.0) for g in gold])) if gold else 0.0


def cmd_analyze(a):
    from vfd.analyst import CONFIGS, Analyst
    from vfd.deep import DeepTool
    from vfd.llm import Chat, pmap
    from vfd.signals import GLiClassScorer

    root = Path(a.bench)
    run = Path(a.run)
    gli = GLiClassScorer(a.gliclass)
    import threading

    lock = threading.Lock()

    class LockedGli:  # one GPU, many analyst threads
        def scores(self, texts, labels):
            with lock:
                return gli.scores(texts, labels)

    deep_lock = threading.Lock()

    class LockedDeep:
        def __init__(self):
            self.d = DeepTool(model=a.deep_model, fast=a.deep_fast, check="lightgbm", max_context=5000)

        def __getattr__(self, name):
            fn = getattr(self.d, name)

            def wrapped(*x, **k):
                with deep_lock:
                    return fn(*x, **k)
            return wrapped

    for cname in a.configs.split(","):
        out = run / f"insights_{cname}.jsonl"
        done = {r["flag"] for r in map(json.loads, out.open())} if out.exists() else set()
        todo = [f for f in flags(root, a.limit) if f.stem not in done]

        def one(fp: Path):
            d = json.load(open(fp))
            df = pd.read_csv(root / d["dataset_csv_path"])
            an = Analyst(Chat(a.model), CONFIGS[cname], deep=LockedDeep(), gli=LockedGli())
            t0 = time.time()
            try:
                r = an.run(df, d["metadata"])
            except Exception as e:
                r = {"config": cname, "insights": [], "summary": "", "error": f"{type(e).__name__}: {str(e)[:200]}"}
            r.update(flag=fp.stem, seconds=round(time.time() - t0, 1), log=an.log[-40:])
            return r

        t0 = time.time()
        rows = pmap(one, todo, workers=a.workers)
        with out.open("a") as f:
            for r in rows:
                f.write(json.dumps(r, default=str) + "\n")
        print(f"[analyze] {cname}: {len(rows)} tables in {time.time() - t0:.0f}s; mean insights "
              f"{np.mean([len(r['insights']) for r in rows]) if rows else 0:.1f}", flush=True)


def cmd_judge(a):
    from vfd.analyst import g_eval_match
    from vfd.llm import Chat, pmap

    root, run = Path(a.bench), Path(a.run)
    judge = Chat(a.model, no_think=True)
    out = run / "scores.jsonl"
    done = {(r["flag"], r["config"]) for r in map(json.loads, out.open())} if out.exists() else set()
    gold = {fp.stem: json.load(open(fp))["insights"] for fp in flags(root, a.limit)}
    items = []
    for f in sorted(run.glob("insights_*.jsonl")):
        for r in map(json.loads, f.open()):
            if (r["flag"], r["config"]) not in done and r["flag"] in gold:
                items.append(r)

    def one(r):
        g = gold[r["flag"]]
        s = g_eval_match(judge, r["insights"], g)
        return {"flag": r["flag"], "config": r["config"], "g_eval": s["g_eval"], "rouge1": rouge1(r["insights"], g),
                "n_pred": len(r["insights"]), "n_gold": len(g), "n_rejected": r.get("n_rejected", 0),
                "llm_calls": r.get("llm_calls", 0), "seconds": r.get("seconds", 0), "error": r.get("error", "")}

    rows = pmap(one, items, workers=a.workers)
    with out.open("a") as f:
        for r in rows:
            f.write(json.dumps(r) + "\n")
    print(f"[judge] scored {len(rows)}")


def cmd_report(a):
    run = Path(a.run)
    df = pd.DataFrame([json.loads(x) for x in (run / "scores.jsonl").open()])
    agg = df.groupby("config").agg(g_eval=("g_eval", "mean"), rouge1=("rouge1", "mean"), tables=("flag", "size"),
                                   insights=("n_pred", "mean"), rejected=("n_rejected", "mean"),
                                   llm_calls=("llm_calls", "mean"), seconds=("seconds", "mean"),
                                   errors=("error", lambda s: int((s != "").sum())))
    base = df[df.config == "D"].set_index("flag")["g_eval"]
    for c in agg.index:
        d = (df[df.config == c].set_index("flag")["g_eval"] - base).dropna()
        agg.loc[c, "vs_D"] = d.mean()
        agg.loc[c, "vs_D_se"] = d.std(ddof=1) / np.sqrt(len(d)) if len(d) > 1 else np.nan
    md = ["# P3: InsightBench (100 tables), analyst agent ablation", "",
          "D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. "
          "g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.",
          "", agg.round(4).to_markdown()]
    (run / "results.md").write_text("\n".join(md) + "\n")
    json.dump(agg.round(5).to_dict("index"), open(run / "metrics.json", "w"), indent=1)
    print("\n".join(md))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("stage", choices=["analyze", "judge", "report"])
    ap.add_argument("--run", required=True)
    ap.add_argument("--bench", default=os.environ.get("INSIGHTBENCH", "artifacts/insight-bench"))
    ap.add_argument("--configs", default="D,E,F")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    ap.add_argument("--deep-model", default="kumo-tabular-l:n_estimators=2,device=cuda")
    ap.add_argument("--deep-fast", default="kumo-tabular-s:n_estimators=1,device=cuda")
    a = ap.parse_args()
    Path(a.run).mkdir(parents=True, exist_ok=True)
    {"analyze": cmd_analyze, "judge": cmd_judge, "report": cmd_report}[a.stage](a)


if __name__ == "__main__":
    main()
