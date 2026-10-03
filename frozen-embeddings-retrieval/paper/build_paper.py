"""Build the paper "Which Layer Is the Embedding?" from the run directories.

  python paper/build_paper.py --roots ../excel_magic/reports/runs

Steps: run scripts/analyze_layers.py (tables + figures into docs/layers), turn its CSVs into the paper's HTML tables,
fill paper/template.html, copy the figures, print paper/frozen-fm-layers-paper.pdf with headless Chromium.
Every number in a table, and the prose numbers written as {{NAME}} placeholders, come from layers_rows.json files.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "scripts"))
import analyze_layers as A  # noqa: E402

CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
ML = A.MODEL_LABEL
TASK_SHORT = {"rel-avito/user-clicks": "avito clicks", "rel-avito/user-visits": "avito visits",
              "rel-event/user-ignore": "event ignore", "rel-event/user-repeat": "event repeat",
              "rel-f1/driver-dnf": "f1 dnf", "rel-f1/driver-top3": "f1 top3", "rel-hm/user-churn": "hm churn",
              "rel-trial/study-outcome": "trial outcome"}
LAYER_SHORT = {"gnn": "graph", "pre_gnn": "row enc.", "row_emb": "row enc.", "final": "final"}


def f3(x):
    return "&mdash;" if x is None or (isinstance(x, float) and not np.isfinite(x)) else f"{x:.3f}"


def lay(name: str) -> str:
    if name in LAYER_SHORT:
        return LAYER_SHORT[name]
    kind, _, i = name.partition("_")
    return f"{'ICL' if kind == 'icl' else 'block'} {int(i)}"


def table_main(P: pd.DataFrame) -> str:
    """J18 Kumo Relational vs J19 TabPFN, real-label context, per task, plus Kumo Relational with random labels."""
    lab = P[P.target == "label"]
    k = lab[lab.study == "J18"].set_index("task")
    t = lab[lab.study == "J19"].set_index("task")
    kr = P[(P.study == "J18") & (P.target == "random")].set_index("task")
    tasks = sorted(set(k.index) & set(t.index))
    rows = []
    for task in tasks:
        a, b, c = k.loc[task], t.loc[task], kr.loc[task] if task in kr.index else None
        best_k = max(a.own, a.last, a.val_pick)
        best_t = max(b.own, b.last, b.val_pick)
        cell = lambda v, best: f"<b>{f3(v)}</b>" if np.isclose(v, best) else f3(v)  # noqa: E731
        rows.append(f"<tr><td>{TASK_SHORT.get(task, task)}</td><td class='n'>{f3(b.raw)}</td>"
                    f"<td class='n'>{cell(a.own, best_k)}</td><td class='n'>{cell(a.last, best_k)}</td>"
                    f"<td class='n'>{cell(a.val_pick, best_k)} <span class='l'>{lay(a.val_layer)}</span></td>"
                    f"<td class='n'>{f3(c.val_pick) if c is not None else '&mdash;'}</td>"
                    f"<td class='n'>{cell(b.own, best_t)}</td><td class='n'>{cell(b.last, best_t)}</td>"
                    f"<td class='n'>{cell(b.val_pick, best_t)} <span class='l'>{lay(b.val_layer)}</span></td></tr>")
    m = lambda d, c: d.loc[tasks, c].mean()  # noqa: E731
    rows.append(f"<tr class='mean'><td>mean ({len(tasks)} tasks)</td><td class='n'>{f3(m(t, 'raw'))}</td>"
                f"<td class='n'>{f3(m(k, 'own'))}</td><td class='n'>{f3(m(k, 'last'))}</td><td class='n'>{f3(m(k, 'val_pick'))}</td>"
                f"<td class='n'>{f3(kr.loc[[x for x in tasks if x in kr.index], 'val_pick'].mean())}</td>"
                f"<td class='n'>{f3(m(t, 'own'))}</td><td class='n'>{f3(m(t, 'last'))}</td><td class='n'>{f3(m(t, 'val_pick'))}</td></tr>")
    head = ("<tr><th rowspan='2'>task</th><th class='n' rowspan='2'>raw<br>features</th>"
            "<th colspan='4' class='c'>Kumo Relational (reads the database)</th><th colspan='3' class='c'>TabPFN v2 (flattened features)</th></tr>"
            "<tr><th class='n'>own</th><th class='n'>last</th><th class='n'>val-picked</th><th class='n'>random ctx</th>"
            "<th class='n'>own</th><th class='n'>last</th><th class='n'>val-picked</th></tr>")
    return head + "".join(rows)


def table_contexts(P: pd.DataFrame) -> str:
    """Label-free vs real-label context, mean over tasks, every model: val-picked layer and last layer."""
    order = [("J18", "kumo-relational"), ("J19", "tabpfn")] + [("J21", m) for m in A.MODEL_ORDER if m != "kumo-relational"]
    targets = ["zeros", "random", "kmeans", "label"]
    rows = []
    for st, mo in order:
        g = P[(P.study == st) & (P.model == mo)]
        if g.empty:
            continue
        n = g.task.nunique()
        raw = g.raw.mean()
        cells = []
        for tg in targets:
            h = g[g.target == tg]
            cells.append(f"<td class='n'>{f3(h.val_pick.mean()) if len(h) else '&mdash;'}</td>")
        own = g[g.target == "label"].own.mean()
        name = ML[mo] + (" (J19)" if st == "J19" else "")
        rows.append(f"<tr><td>{name}</td><td class='n'>{n}</td><td class='n'>{f3(raw)}</td>{''.join(cells)}<td class='n'>{f3(own)}</td></tr>")
    head = ("<tr><th>model</th><th class='n'>tasks</th><th class='n'>raw</th><th class='n'>zeros</th><th class='n'>random</th>"
            "<th class='n'>k-means</th><th class='n'>real labels</th><th class='n'>own pred.</th></tr>")
    return head + "".join(rows)


def table_selectors(R: pd.DataFrame) -> str:
    if R.empty:
        return ""
    sel = [("val", "validation labels"), ("last", "last layer"), ("random layer", "average layer")]
    sel += [(f"{s}:loto", A.GEO_LABEL[s]) for s in A.GEO if f"{s}:loto" in R]
    rows = []
    for key, name in sorted(sel, key=lambda kv: R[kv[0]].mean()):
        rho = R.get(f"{key.split(':')[0]}:rho")
        rho_s = f"{rho.mean():+.2f}" if rho is not None else "&mdash;"
        rows.append(f"<tr><td>{name}</td><td class='n'>{f3(R[key].mean())}</td><td class='n'>{f3(R[key].median())}</td>"
                    f"<td class='n'>{rho_s}</td></tr>")
    head = "<tr><th>layer chosen by</th><th class='n'>mean regret</th><th class='n'>median</th><th class='n'>Spearman &rho;</th></tr>"
    return head + "".join(rows)


def table_protocol(P: pd.DataFrame) -> str:
    K = ["study", "model", "task", "target"]
    a, b = P[P.split == "random"].set_index(K), P[P.split == "time"].set_index(K)
    both = a.index.intersection(b.index)
    rows = []
    for tg in ["random", "kmeans", "label"]:
        idx = [i for i in both if i[3] == tg]
        if not idx:
            continue
        for name, X in (("random", a), ("time", b)):
            x = X.loc[idx]
            rows.append(f"<tr><td>{A.TARGET_LABEL[tg]}</td><td>{name}</td><td class='n'>{len(idx)}</td>"
                        f"<td class='n'>{f3(x['last'].mean())}</td><td class='n'>{f3(x['last_knn'].mean())}</td>"
                        f"<td class='n'>{f3(x['val_pick'].mean())}</td></tr>")
    head = ("<tr><th>context</th><th>split</th><th class='n'>runs</th><th class='n'>last, linear</th>"
            "<th class='n'>last, kNN</th><th class='n'>val-picked</th></tr>")
    return head + "".join(rows)


def cka_summary(j21: list[dict]) -> dict:
    """Mean cross-model linear CKA between best layers, between last layers, and to the raw features."""
    mats: dict = {}
    for r in j21:
        c = r.get("cka_models", {}).get("label")
        if not c:
            continue
        for i, a in enumerate(c["keys"]):
            for j, b in enumerate(c["keys"]):
                mats.setdefault((a, b), []).append(c["matrix"][i][j])
    M = {k: float(np.mean(v)) for k, v in mats.items()}
    models = sorted({k.split(":")[0] for k, _ in M if ":" in k})
    sdm = [m for m in models if m.startswith(("kumo", "tabicl"))]
    out = {}
    for kind in ("best", "last"):
        pairs = [M[(f"{a}:{kind}", f"{b}:{kind}")] for i, a in enumerate(models) for b in models[i + 1:]
                 if (f"{a}:{kind}", f"{b}:{kind}") in M]
        fam = [M[(f"{a}:{kind}", f"{b}:{kind}")] for i, a in enumerate(sdm) for b in sdm[i + 1:]
               if (f"{a}:{kind}", f"{b}:{kind}") in M]
        raw = [M[(f"{a}:{kind}", "raw_features")] for a in models if (f"{a}:{kind}", "raw_features") in M]
        out[kind] = {"all_pairs": np.mean(pairs) if pairs else np.nan, "sdm_pairs": np.mean(fam) if fam else np.nan,
                     "to_raw": np.mean(raw) if raw else np.nan}
    out["n_tasks"] = len(j21)
    return out


def table_cka(c: dict) -> str:
    rows = [f"<tr><td>{k} layer</td><td class='n'>{c[k]['all_pairs']:.2f}</td><td class='n'>{c[k]['sdm_pairs']:.2f}</td>"
            f"<td class='n'>{c[k]['to_raw']:.2f}</td></tr>" for k in ("best", "last")]
    head = ("<tr><th></th><th class='n'>all model pairs</th><th class='n'>TabICL/Kumo family</th>"
            "<th class='n'>vs raw features</th></tr>")
    return head + "".join(rows)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--roots", nargs="+", default=[str(ROOT.parent / "excel_magic" / "reports" / "runs")])
    a = ap.parse_args()
    layers_dir = ROOT / "docs" / "layers"
    subprocess.run([sys.executable, str(ROOT / "scripts" / "analyze_layers.py"), "--roots", *a.roots,
                    "--out", str(layers_dir)], check=True, stdout=subprocess.DEVNULL)
    P = pd.read_csv(layers_dir / "layers_per_run.csv")
    R = pd.read_csv(layers_dir / "layers_selectors.csv") if (layers_dir / "layers_selectors.csv").exists() else pd.DataFrame()
    _, _, j21_all = A.load(a.roots)
    j21 = [r for r in j21_all if r.get("ctx_split", "random") == "time"]
    Pt = P[P.split == "time"]
    c = cka_summary(j21)
    n21 = len(j21)
    fills = {"T_MAIN": table_main(Pt), "T_CONTEXTS": table_contexts(Pt), "T_SELECTORS": table_selectors(R),
             "T_PROTOCOL": table_protocol(P), "T_CKA": table_cka(c), "N_J21": str(n21),
             "J21_TASKS": ", ".join(TASK_SHORT.get(f"{r['dataset']}/{r['task']}", r["task"]) for r in j21),
             "CKA_BEST": f"{c['best']['all_pairs']:.2f}", "CKA_LAST": f"{c['last']['all_pairs']:.2f}",
             "CKA_RAW_BEST": f"{c['best']['to_raw']:.2f}", "CKA_RAW_LAST": f"{c['last']['to_raw']:.2f}"}
    if not R.empty:
        fills.update({"REG_VAL": f"{R['val'].mean():.3f}", "REG_LAST": f"{R['last'].mean():.3f}",
                      "REG_ID": f"{R['intrinsic_dim:loto'].mean():.3f}", "REG_ER": f"{R['effective_rank:loto'].mean():.3f}",
                      "RHO_ER": f"{R['effective_rank:rho'].mean():+.2f}",
                      "ER_NEG": f"{(R['effective_rank:rho'] < 0).mean():.0%}".replace("%", "&thinsp;%"),
                      "N_SEL": str(R['val'].notna().sum())})
    lab = Pt[Pt.target == "label"]
    k18, t19 = lab[lab.study == "J18"], lab[lab.study == "J19"]
    r18 = Pt[(Pt.study == "J18") & (Pt.target == "random")]
    proto = P.set_index(["study", "model", "task", "target"])
    both = proto[proto.split == "random"].index.intersection(proto[proto.split == "time"].index)
    lab_both = [i for i in both if i[3] == "label"]
    fills.update({"K_VAL": f3(k18.val_pick.mean()), "K_OWN": f3(k18.own.mean()), "K_LAST": f3(k18.last.mean()),
                  "K_WIN_OWN": str(int((k18.val_pick > k18.own).sum())), "K_N": str(len(k18)),
                  "K_RAND": f3(r18.val_pick.mean()), "RAW": f3(t19.raw.mean()),
                  "T_VAL": f3(t19.val_pick.mean()), "T_OWN": f3(t19.own.mean()), "T_LAST": f3(t19.last.mean()),
                  "T_B89": str(int(t19.val_layer.isin(["block_08", "block_09"]).sum())),
                  "P_LAST_RAND": f3(proto[proto.split == "random"].loc[lab_both, "last"].mean()),
                  "P_LAST_TIME": f3(proto[proto.split == "time"].loc[lab_both, "last"].mean()),
                  "P_N": str(len(lab_both))})
    html = (HERE / "template.html").read_text()
    for k, v in fills.items():
        html = html.replace("{{" + k + "}}", v)
    assert "{{" not in html, "unfilled placeholder"
    (HERE / "paper.html").write_text(html)
    fig = HERE / "figures"
    fig.mkdir(exist_ok=True)
    for f in ("fig_layers_depth.png", "fig_layers_geometry.png", "fig_layers_selectors.png", "fig_layers_cka.png"):
        if (layers_dir / f).exists():
            shutil.copy2(layers_dir / f, fig / f)
    subprocess.run([CHROME, "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={HERE / 'frozen-fm-layers-paper.pdf'}", f"file://{HERE / 'paper.html'}"],
                   check=True, stderr=subprocess.DEVNULL)
    print("J21 tasks:", n21, "| CKA:", {k: v for k, v in c.items()})
    print("selectors:", R[["val", "last"] + [f"{s}:loto" for s in A.GEO if f"{s}:loto" in R]].mean().round(3).to_dict() if not R.empty else "-")


if __name__ == "__main__":
    main()
