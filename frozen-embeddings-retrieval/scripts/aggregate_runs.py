"""Aggregate runs/*bench_* metrics into docs/run-log.md (mean +/- sd over seeds) and docs/results_long.csv."""
import glob
import json
import sys

import numpy as np
import pandas as pd

rows = []
used = []
for d in sorted(set(glob.glob("runs/*bench_*") + glob.glob("runs/*retrieval*") + glob.glob("runs/*exp04*"))):
    import os
    if not os.path.exists(f"{d}/metrics.json"):
        print("skipping incomplete run", d); continue
    if os.path.exists(f"{d}/rescored.json"):  # preferred: scored with the current benchmark code from saved embeddings
        rs = json.load(open(f"{d}/rescored.json"))
        results = [{"embedder": k, "status": "ok", "seconds": np.nan, **v} for k, v in rs["results"].items()]
        dbname, seed = rs["db"], rs["seed"]
    else:
        m = json.load(open(f"{d}/metrics.json"))
        if m.get("status") != "ok" or "results" not in m:
            print("skipping failed run", d); continue
        results, dbname = m["results"], m["db"]
        seed = json.load(open(f"{d}/config.json"))["config"]["seed"]
    used.append(d.split("/")[-1])
    secs = {}
    try:
        secs = {r["embedder"]: r.get("seconds", np.nan) for r in json.load(open(f"{d}/metrics.json")).get("results", [])}
    except Exception:
        pass
    for r in results:
        if r["status"] != "ok":
            print("ERR", d, r["embedder"], r.get("error")); continue
        s = r.get("segment", {}); f = r["future_purchase"]; fn = r.get("future_purchase_novel", {})
        rows.append({"db": dbname, "seed": seed, "embedder": r["embedder"], "seg_P@10": s.get("P@10", np.nan),
                     "seg_kNN_acc": s.get("kNN_acc", np.nan), "fut_MAP@10": f["MAP@10"], "fut_recall@10": f["recall@10"],
                     "pop_MAP@10": f["pop_MAP@10"], "novel_MAP@10": fn.get("MAP@10", np.nan),
                     "novel_pop_MAP@10": fn.get("pop_MAP@10", np.nan), "novel_recall@10": fn.get("recall@10", np.nan),
                     "n_queries": f["n_queries"], "seconds": secs.get(r["embedder"], r.get("seconds", np.nan))})
df = pd.DataFrame(rows).drop_duplicates(subset=["db", "seed", "embedder"], keep="last")
order = ["row", "agg", "tabpfn_row", "tabpfn_agg", "tabpfn_agg_rand4", "tabpfn_agg_kmeans", "tabpfn_agg_feat4", "tabpfn_agg_churn", "gnn", "openrfm", "openrfm_ctx16", "openrfm_random", "kumo_relational"]
lines = []
for db, g in df.groupby("db"):
    agg = g.groupby("embedder").agg(seg_P10=("seg_P@10", "mean"), seg_P10_sd=("seg_P@10", "std"), knn=("seg_kNN_acc", "mean"),
                                     fut=("fut_MAP@10", "mean"), fut_sd=("fut_MAP@10", "std"), rec=("fut_recall@10", "mean"),
                                     popmap=("pop_MAP@10", "mean"), nov=("novel_MAP@10", "mean"), nov_sd=("novel_MAP@10", "std"),
                                     novpop=("novel_pop_MAP@10", "mean"), n=("seed", "count"), sec=("seconds", "mean"), nq=("n_queries", "max"))
    agg = agg.reindex([o for o in order if o in agg.index] + [e for e in agg.index if e not in order])
    lines.append(f"\n### {db} ({int(agg.n.max())} seed(s), {int(agg.nq.max())} query entities; popularity baseline MAP@10 = {agg.popmap.iloc[0]:.3f}, novel-only popularity = {agg.novpop.iloc[0]:.3f})\n")
    lines.append("| embedder | seg P@10 (mean +/- sd) | seg 10-NN acc | fut MAP@10 (mean +/- sd) | fut recall@10 | novel MAP@10 (mean +/- sd) | s |")
    lines.append("|---|---|---|---|---|---|---|")
    for e, r in agg.iterrows():
        sd = 0 if np.isnan(r.seg_P10_sd) else r.seg_P10_sd
        segp = "-" if np.isnan(r.seg_P10) else f"{r.seg_P10:.3f} +/- {sd:.3f}"
        knn = "-" if np.isnan(r.knn) else f"{r.knn:.3f}"
        fsd = 0 if np.isnan(r.fut_sd) else r.fut_sd
        nsd = 0 if np.isnan(r.nov_sd) else r.nov_sd
        sec = "-" if np.isnan(r.sec) else f"{r.sec:.0f}"
        lines.append(f"| {e} | {segp} | {knn} | {r.fut:.3f} +/- {fsd:.3f} | {r.rec:.3f} | {r.nov:.3f} +/- {nsd:.3f} | {sec} |")
txt = "\n".join(lines)
print(txt)
if "--write" in sys.argv:
    open("docs/run-log.md", "a").write(f"\n## Aggregated (runs: {', '.join(used)})\n{txt}\n")
    df.to_csv("docs/results_long.csv", index=False)
