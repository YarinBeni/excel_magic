"""P1 / claim C5: do GLiClass labels of free-text columns help a frozen tabular foundation model more than dense text
features? 18 datasets of the AutoGluon multimodal text+tabular benchmark (Shi et al., NeurIPS 2021), official splits.

  prepare   (vfd venv, CPU)  download, subsample train to <= MAX_TRAIN and test to <= MAX_TEST rows, save parquet
  labels    (vfd venv, vLLM) LLM label sets per text column: task-agnostic and task-aware
  features  (vfd venv, GPU)  GLiClass scores, tf-idf+SVD, bge-small+PCA per text column
  evaluate  (tabfm venv, GPU) variants x models: Kumo Tabular-L, TabPFN-2.5, LightGBM; official metric per dataset
  report
Variants: base (numeric + categorical only), tfidf, embed, gli_agnostic, gli_task, gli_task+embed.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

ALIASES = ["prod", "airbnb", "channel", "wine", "imdb", "jigsaw", "fake", "kick", "ae", "qaa", "qaq", "cloth",
           "mercari", "jc", "pop", "book", "salary", "house"]
VARIANTS = ["base", "tfidf", "embed", "gli_agnostic", "gli_task", "gli_task_embed"]


def cmd_prepare(a):
    from auto_mm_bench.datasets import TEXT_BENCHMARK_ALIAS_MAPPING, create_dataset

    out = Path(a.data)
    out.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(0)
    for al in a.datasets.split(","):
        name = TEXT_BENCHMARK_ALIAS_MAPPING[al]
        if (out / name / "meta.json").exists():
            continue
        tr, te = create_dataset(name, "train"), create_dataset(name, "test")
        types = dict(zip(tr.feature_columns, tr.feature_types))
        meta = {"alias": al, "name": name, "label": tr.label_columns[0], "problem_type": tr.problem_type,
                "metric": tr.metric, "text": [c for c, t in types.items() if t == "text"],
                "categorical": [c for c, t in types.items() if t == "categorical"],
                "numerical": [c for c, t in types.items() if t == "numerical"],
                "n_train_full": len(tr.data), "n_test_full": len(te.data)}
        dtr, dte = tr.data, te.data
        if len(dtr) > a.max_train:
            dtr = dtr.iloc[np.sort(rng.choice(len(dtr), a.max_train, replace=False))]
        if len(dte) > a.max_test:
            dte = dte.iloc[np.sort(rng.choice(len(dte), a.max_test, replace=False))]
        (out / name).mkdir(exist_ok=True)
        cols = meta["text"] + meta["categorical"] + meta["numerical"] + [meta["label"]]
        dtr[cols].reset_index(drop=True).to_parquet(out / name / "train.parquet")
        dte[cols].reset_index(drop=True).to_parquet(out / name / "test.parquet")
        meta.update(n_train=len(dtr), n_test=len(dte))
        json.dump(meta, open(out / name / "meta.json", "w"), indent=1)
        print("[prepare]", al, name, meta["problem_type"], meta["metric"], "text", meta["text"], len(dtr), len(dte))


def datasets(a):
    for d in sorted(Path(a.data).iterdir()):
        if (d / "meta.json").exists():
            yield d, json.load(open(d / "meta.json"))


def cmd_labels(a):
    from vfd.featurize import propose_labels
    from vfd.llm import Chat

    chat = Chat(a.model)
    for d, m in datasets(a):
        if (d / "labels.json").exists():
            continue
        tr = pd.read_parquet(d / "train.parquet")
        out = {}
        for c in m["text"]:
            s = tr[c].dropna().astype(str)
            samples = s.sample(min(20, len(s)), random_state=0).tolist()
            out[c] = {"agnostic": propose_labels(chat, c, samples, k=12),
                      "task": propose_labels(chat, c, samples, k=12, target=m["label"])}
        json.dump(out, open(d / "labels.json", "w"), indent=1)
        print("[labels]", m["alias"], {c: v["task"][:4] for c, v in out.items()})


def cmd_features(a):
    from sentence_transformers import SentenceTransformer

    from vfd.featurize import embed_pca, gli_features, tfidf_svd
    from vfd.signals import GLiClassScorer

    gli = GLiClassScorer(a.gliclass, batch_size=32)
    emb = SentenceTransformer("BAAI/bge-small-en-v1.5", device="cuda")
    lat = {}
    for d, m in datasets(a):
        if (d / "features_train.parquet").exists():
            continue
        tr, te = pd.read_parquet(d / "train.parquet"), pd.read_parquet(d / "test.parquet")
        labels = json.load(open(d / "labels.json"))
        ftr, fte = {}, {}
        t0 = time.time()
        for c in m["text"]:
            for kind in ("agnostic", "task"):
                L = labels[c][kind]
                a_tr, a_te = gli_features(gli, tr[c].tolist(), L), gli_features(gli, te[c].tolist(), L)
                for j, lab in enumerate(L):
                    ftr[f"gli_{kind}__{c}__{j}"], fte[f"gli_{kind}__{c}__{j}"] = a_tr[:, j], a_te[:, j]
            x_tr, x_te = tfidf_svd(tr[c].tolist(), te[c].tolist(), 32)
            for j in range(x_tr.shape[1]):
                ftr[f"tfidf__{c}__{j}"], fte[f"tfidf__{c}__{j}"] = x_tr[:, j], x_te[:, j]
            e_tr, e_te = embed_pca(emb, tr[c].tolist(), te[c].tolist(), 16)
            for j in range(e_tr.shape[1]):
                ftr[f"embed__{c}__{j}"], fte[f"embed__{c}__{j}"] = e_tr[:, j], e_te[:, j]
        pd.DataFrame(ftr).to_parquet(d / "features_train.parquet")
        pd.DataFrame(fte).to_parquet(d / "features_test.parquet")
        lat[m["alias"]] = round(time.time() - t0, 1)
        print("[features]", m["alias"], len(ftr), "columns", lat[m["alias"]], "s")
    json.dump(lat, open(Path(a.data) / "features_seconds.json", "w"), indent=1)


def variant_frame(df: pd.DataFrame, feats: pd.DataFrame, m: dict, variant: str) -> pd.DataFrame:
    base = df[m["categorical"] + m["numerical"]].copy()
    for c in m["categorical"]:
        base[c] = base[c].astype(str)
    pick = {"base": [], "tfidf": ["tfidf__"], "embed": ["embed__"], "gli_agnostic": ["gli_agnostic__"],
            "gli_task": ["gli_task__"], "gli_task_embed": ["gli_task__", "embed__"]}[variant]
    cols = [c for c in feats.columns if any(c.startswith(p) for p in pick)]
    return pd.concat([base, feats[cols]], axis=1) if cols else base


def score(metric: str, y, pred, proba, classes) -> float:
    from sklearn.metrics import accuracy_score, r2_score, roc_auc_score

    if metric == "acc":
        return float(accuracy_score(y, pred))
    if metric == "roc_auc":
        classes = list(classes)
        pos = classes.index(1) if 1 in classes else len(classes) - 1  # label 1 if present, else the last class
        return float(roc_auc_score(np.asarray(y) == classes[pos], proba[:, pos]))
    if metric == "r2":
        return float(r2_score(y, pred))
    raise ValueError(metric)


def cmd_evaluate(a):
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "open-tabfm-auto"))
    from tabfm_auto.models.registry import get_model

    from vfd.tabular import encode_frames, encode_target

    rows = []
    out = Path(a.run) / "rows.jsonl"
    done = set()
    if out.exists():
        done = {(r["dataset"], r["variant"], r["model"]) for r in map(json.loads, out.open())}
    for d, m in datasets(a):
        tr, te = pd.read_parquet(d / "train.parquet"), pd.read_parquet(d / "test.parquet")
        ftr, fte = pd.read_parquet(d / "features_train.parquet"), pd.read_parquet(d / "features_test.parquet")
        task = {"binary": "binary", "multiclass": "multiclass", "regression": "regression"}[m["problem_type"]]
        if task == "regression":
            y_tr, y_te, classes = tr[m["label"]].astype(float).to_numpy(), te[m["label"]].astype(float).to_numpy(), None
        else:
            (y_tr, y_te), classes = encode_target(tr[m["label"]], te[m["label"]])
        for variant in VARIANTS:
            Xtr, Xte = variant_frame(tr, ftr, m, variant), variant_frame(te, fte, m, variant)
            if Xtr.shape[1] == 0:
                continue
            Xtr, Xte = encode_frames(Xtr, Xte)
            for spec in a.models.split(";"):
                key = (m["alias"], variant, spec)
                if key in done:
                    continue
                t0 = time.time()
                try:
                    est = get_model(spec, task)
                    est.fit(Xtr, y_tr)
                    pred = est.predict(Xte)
                    proba = est.predict_proba(Xte) if task != "regression" else None
                    val = score(m["metric"], y_te, pred, proba, list(getattr(est, "classes_", np.unique(y_tr))))
                    err = ""
                except Exception as e:  # report and continue: one bad (dataset, model) must not stop the sweep
                    val, err = float("nan"), f"{type(e).__name__}: {str(e)[:200]}"
                r = {"dataset": m["alias"], "variant": variant, "model": spec, "metric": m["metric"], "value": val,
                     "n_features": int(Xtr.shape[1]), "seconds": round(time.time() - t0, 1), "error": err}
                rows.append(r)
                with out.open("a") as f:
                    f.write(json.dumps(r) + "\n")
                print("[evaluate]", r, flush=True)


def cmd_report(a):
    run = Path(a.run)
    df = pd.DataFrame([json.loads(x) for x in (run / "rows.jsonl").open()])
    df["model"] = df["model"].str.split(":").str[0]
    piv = df.pivot_table(index=["dataset", "metric"], columns=["model", "variant"], values="value")
    md = ["# P1: GLiClass labels of text columns as features for frozen tabular models", "",
          "Official test metric per dataset (acc, roc_auc or r2; higher is better).", "", piv.round(4).to_markdown(), ""]
    # relative to base, per model: mean gain and wins
    md += ["| model | variant | mean change vs base | datasets better than base | datasets better than embed |",
           "|---|---|---|---|---|"]
    for mo in sorted(df.model.unique()):
        p = df[df.model == mo].pivot_table(index="dataset", columns="variant", values="value")
        for v in VARIANTS[1:]:
            if v not in p or "base" not in p:
                continue
            d = (p[v] - p["base"]).dropna()
            vs_e = (p[v] - p["embed"]).dropna() if "embed" in p and v != "embed" else pd.Series(dtype=float)
            md.append(f"| {mo} | {v} | {d.mean():+.4f} | {(d > 0).sum()}/{len(d)} | "
                      f"{(vs_e > 0).sum()}/{len(vs_e) if len(vs_e) else '-'} |")
    (run / "results.md").write_text("\n".join(md) + "\n")
    df.to_csv(run / "results.csv", index=False)
    json.dump({"n_rows": len(df), "errors": int((df.error != "").sum())}, open(run / "metrics.json", "w"))
    print("\n".join(md))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("stage", choices=["prepare", "labels", "features", "evaluate", "report"])
    ap.add_argument("--data", default=os.environ.get("MMBENCH_DATA", "artifacts/mmbench"))
    ap.add_argument("--run", default="runs/p1")
    ap.add_argument("--datasets", default=",".join(ALIASES))
    ap.add_argument("--max-train", type=int, default=3000)
    ap.add_argument("--max-test", type=int, default=2000)
    ap.add_argument("--model", default="Qwen/Qwen3-Coder-30B-A3B-Instruct")
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    ap.add_argument("--models", default="kumo-tabular-l:n_estimators=4,device=cuda;tabpfn-2.5:n_estimators=4,device=cuda;lightgbm")
    a = ap.parse_args()
    Path(a.run).mkdir(parents=True, exist_ok=True)
    {"prepare": cmd_prepare, "labels": cmd_labels, "features": cmd_features, "evaluate": cmd_evaluate,
     "report": cmd_report}[a.stage](a)


if __name__ == "__main__":
    main()
