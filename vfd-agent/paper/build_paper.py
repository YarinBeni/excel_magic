"""Build paper 3, "Verified, Fast, Deep?", from the run directories.

  python vfd-agent/paper/build_paper.py --runs reports/runs

Every table and every prose number written as {{NAME}} in template.html is computed here from the run files
(metrics.json, rows / scores / capability jsonl). The PDF is printed with headless Chromium.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"


def latest(runs: Path, pattern: str, need: str) -> Path | None:
    ds = sorted(d for d in runs.glob(f"*{pattern}") if (d / need).exists())
    return ds[-1] if ds else None


def jl(p: Path) -> list[dict]:
    return [json.loads(x) for x in p.open()] if p.exists() else []


def f3(x) -> str:
    return "&mdash;" if x is None or not np.isfinite(x) else f"{x:.3f}"


def sg(x) -> str:
    return "&mdash;" if x is None or not np.isfinite(x) else f"{x:+.3f}"


def paired(a: pd.Series, b: pd.Series) -> tuple[float, float]:
    d = (a - b).dropna()
    return (float(d.mean()), float(d.std(ddof=1) / np.sqrt(len(d)))) if len(d) > 1 else (np.nan, np.nan)


def table(head: list[str], rows: list[list], caption: str, cls: str = "", num_from: int = 1) -> str:
    th = "".join(f'<th class="n">{h}</th>' if i >= num_from else f"<th>{h}</th>" for i, h in enumerate(head))
    trs = "".join("<tr>" + "".join(f'<td class="n">{c}</td>' if i >= num_from else f"<td>{c}</td>"
                                   for i, c in enumerate(r)) + "</tr>" for r in rows)
    return f'<table class="{cls}"><caption>{caption}</caption><tr>{th}</tr>{trs}</table>'


# ------------------------------------------------------------------------------------------------- P0 verifiers
SIG = [("sig_exec_ok", "query executes"), ("sig_nonempty", "result not empty"), ("sig_logprob", "generator log-prob"),
       ("sig_self_consistency", "self-consistency (vote share)"), ("sig_gliclass", "<b>GLiClass, zero-shot</b>"),
       ("sig_nli", "NLI cross-encoder"), ("sig_judge", "LLM judge (other family)"),
       ("stack_cheap", "stack: no-model signals"), ("stack_cheap_gli", "<b>stack: + GLiClass</b>"),
       ("stack_all", "stack: + NLI + judge")]
LINK = [("gliclass_large_v3", "<b>GLiClass large v3</b>"), ("lexical", "lexical overlap"),
        ("rerank_bge_m3", "bge-reranker-v2-m3"), ("embed_bge_small", "bge-small embedding"),
        ("gliner_bi_base_v2", "GLiNER bi-encoder")]


def p0(runs: Path, V: dict) -> dict:
    d = latest(runs, "V1_verifier_study_all", "metrics.json")
    m = json.load(open(d / "metrics.json"))
    s, lat = m["signals"], m["latency"]
    ms = {"sig_gliclass": lat["gliclass_s_per_item"] * 1000, "sig_nli": lat["nli_s_per_item"] * 1000,
          "sig_judge": lat["judge_s_per_call_sequential"] * 1000}
    rows = [[lab, f3(s[k]["auroc"]), f3(s[k]["within_question_auroc"]), f3(s[k]["best_of_n"]),
             f"{ms[k]:.0f}" if k in ms else ("0" if k.startswith("sig") else "")] for k, lab in SIG if k in s]
    T = {"T_P0": table(["signal", "AUROC", "within-question AUROC", "best-of-N acc.", "ms / item"], rows,
                       f"Table 2. Which signal tells a correct SQL candidate from a wrong one? BIRD Arcwise-Plat, "
                       f"{m['n_questions']} questions, {m['n_candidates']:,} candidates (1 greedy + 8 sampled). "
                       f"Greedy accuracy {m['greedy_accuracy']:.3f}, oracle (any candidate correct) "
                       f"{m['oracle_accuracy']:.3f}. Stacks: logistic regression, 2-fold by question.")}
    sl = m["schema_link"]
    T["T_LINK"] = table(["linker", "@5", "@10", "@20", "all@20", "ms"],
                        [[lab, f3(sl[k]["recall@5"]), f3(sl[k]["recall@10"]), f3(sl[k]["recall@20"]),
                          f3(sl[k]["all@20"]), f"{lat[k] * 1000:.0f}"] for k, lab in LINK],
                        "Table 3. Schema linking (497 questions): share of the gold SQL's columns in the linker's top k; all@20: "
                        "questions with every gold column in the top 20; ms per question.")
    V.update(P0_N=m["n_questions"], P0_C=f"{m['n_candidates']:,}", P0_GREEDY=f3(m["greedy_accuracy"]),
             P0_ORACLE=f3(m["oracle_accuracy"]), P0_GLI=f3(s["sig_gliclass"]["auroc"]),
             P0_SC=f3(s["sig_self_consistency"]["auroc"]), P0_LP=f3(s["sig_logprob"]["auroc"]),
             P0_EXEC=f3(s["sig_exec_ok"]["auroc"]), P0_NLI=f3(s["sig_nli"]["auroc"]),
             P0_JUDGE=f3(s["sig_judge"]["auroc"]), P0_STACK=f3(s["stack_cheap"]["auroc"]),
             P0_STACKGLI=f3(s["stack_cheap_gli"]["auroc"]), P0_STACKALL=f3(s["stack_all"]["auroc"]),
             P0_GLI_MS=f"{ms['sig_gliclass']:.0f}", P0_JUDGE_MS=f"{ms['sig_judge']:.0f}",
             P0_SPEED=f"{ms['sig_judge'] / ms['sig_gliclass']:.0f}",
             L_GLI10=f3(sl["gliclass_large_v3"]["recall@10"]), L_LEX10=f3(sl["lexical"]["recall@10"]),
             L_GLINER10=f3(sl["gliner_bi_base_v2"]["recall@10"]), L_GLI_ALL20=f3(sl["gliclass_large_v3"]["all@20"]),
             L_MISS20=f"{100 * (1 - sl['gliclass_large_v3']['all@20']):.0f}")
    return T


# ------------------------------------------------------------------------------------------------- P1 text features
def p1(runs: Path, V: dict) -> dict:
    d = latest(runs, "V2_text_featurizer", "rows.jsonl")
    r = pd.DataFrame(jl(d / "rows.jsonl"))
    r["m"] = r.model.str.split(":").str[0]
    p = r.pivot_table(index=["dataset", "m"], columns="variant", values="value")
    rows = []
    for m, lab in [("kumo-tabular-l", "Kumo Tabular-L"), ("lightgbm", "LightGBM"), ("tabpfn-2.5", "TabPFN-2.5")]:
        q = p.xs(m, level=1)
        cell = []
        for v in ["tfidf", "gli_task", "embed", "gli_task_embed"]:
            x = (q[v] - q["base"]).dropna()
            cell.append(f"{x.mean():+.3f} ({(x > 0).sum()}/{len(x)})")
        both = (q["gli_task_embed"] - q["embed"]).dropna()
        cell.append(f"{(both > 0).sum()}/{len(both)}")
        rows.append([lab] + cell)
        V[f"P1_{m.split('-')[0].upper()}_BEAT"] = f"{(both > 0).sum()} of {len(both)}"
    V["P1_N"] = r.dataset.nunique()
    return {"T_P1": table(["model", "tf-idf + SVD", "GLiClass labels", "embedding", "labels + embedding",
                           "labels + emb. beat emb."], rows,
                          f"Table 6. Text columns as features. Mean change of the official test metric (accuracy, AUROC "
                          f"or R<sup>2</sup>) against numeric and categorical columns only; in brackets, datasets that "
                          f"improve. {V['P1_N']} datasets of the AutoGluon multimodal benchmark.", "wide")}


# ------------------------------------------------------------------------------------------------- P2 + P4 SQL harness
P2 = [("A", "A: LLM alone, greedy, full schema"), ("B", "B: + GLiClass schema linking (top 20 + keys)"),
      ("C", "C: + GLiClass triage, one revision"), ("D", "D: + 8 candidates, cheap verifier stack"),
      ("F", "F: D + LLM judge on the uncertain band"), ("SC", "SC: 8 candidates, majority vote, no small models")]


def p2(runs: Path, V: dict) -> dict:
    d = latest(runs, "V5_sql_harness", "metrics.json")
    m = json.load(open(d / "metrics.json"))
    rows = [[lab, f3(m[k]["accuracy"]), "" if k == "A" else f"{m[k]['vs_A']:+.3f} ({m[k]['vs_A_se']:.3f})",
             f"{m[k]['schema_chars']:,.0f}", f"{m[k]['llm_calls']:.2f}" + (f" + {m[k]['judge_calls']:.1f}"
                                                                          if m[k]["judge_calls"] else ""),
             f"{m[k]['seconds']:.1f}"] for k, lab in P2 if k in m]
    T = {"T_P2": table(["configuration", "accuracy", "vs A (se)", "schema chars", "LLM calls", "s / question"], rows,
                       f"Table 4. The SQL harness, end to end. BIRD Arcwise-Plat, {m['A']['n']} questions, "
                       f"Qwen3-Coder-30B-A3B. Paired difference against A, standard error over questions.", "wide")}
    V.update({f"P2_{k}": f3(m[k]["accuracy"]) for k, _ in P2 if k in m})
    V.update(P2_B_D=f"{m['B']['vs_A']:+.3f}", P2_B_SE=f"{m['B']['vs_A_se']:.3f}", P2_N=m["A"]["n"],
             P2_CH_A=f"{m['A']['schema_chars']:,.0f}", P2_CH_B=f"{m['B']['schema_chars']:,.0f}")
    d = latest(runs, "V5_autoresearch_sql", "metrics.json")
    a = json.load(open(d / "metrics.json"))
    hr = []
    for h in a["history"]:
        hyp = h["hypothesis"] if not h["hypothesis"].startswith("(fallback") else "(no proposal; untried value)"
        hr.append([h["i"] + 1, f"<code>{h['knob']}={h['value']}</code>", f"{h['search_delta']:+.3f} ({h['search_se']:.3f})",
                   f"{h['accept_delta']:+.3f}", "<b>kept</b>" if h["kept"] else "rejected"])
    T["T_P4"] = table(["#", "change", "search gain (se)", "accept gain", "decision"], hr,
                      f"Table 5. The guarded auto-research loop. Researcher GLM-4.5-Air; split fixed before the loop: "
                      f"search {a['n_search']}, accept {a['n_accept']}, held-out {a['n_heldout']} questions. Keep "
                      f"rule: search gain &gt; 1 se and accept gain &ge; 0.", "wide", num_from=2)
    kept = [h for h in a["history"] if h["kept"]]
    V.update(P4_S=a["n_search"], P4_A=a["n_accept"], P4_H=a["n_heldout"], P4_ATT=a["attempts"], P4_KEPT=a["kept"],
             P4_REJ=a["attempts"] - a["kept"], P4_H0=f3(a["heldout_initial"]), P4_H1=f3(a["heldout_final"]),
             P4_HD=f"{a['heldout_delta']:+.3f}", P4_HSE=f"{a['heldout_se']:.3f}",
             P4_KEPT_WHAT=", ".join(f"<code>{h['knob']}={h['value']}</code>" for h in kept) or "nothing")
    noop = [h for h in a["history"] if h["knob"] == "link_k" and a["initial"]["link"] is None]
    V["P4_NOISE"] = f"{max(abs(h['search_delta']) for h in noop):.3f}" if noop else "&mdash;"
    return T


# ------------------------------------------------------------------------------------------------- P3 InsightBench
def p3(runs: Path, V: dict) -> dict:
    sc = []
    for d in sorted(runs.glob("*V3_insightbench_all")) + sorted(runs.glob("*V6_insightbench_pi_all")):
        sc += [dict(x, run=d.name) for x in jl(d / "scores.jsonl")]
    s = pd.DataFrame(sc)
    s = s.drop_duplicates(["flag", "config"], keep="last")          # the later run of a config replaces the earlier
    piv = s.pivot(index="flag", columns="config", values="g_eval")
    lab = {"D": "D: LLM + SQL", "E": "E: + DEEP tools (tabular FM)", "F": "F: + text labels + verified ledger",
           "piD": "pi + D tools", "piE": "pi + E tools", "piF": "pi + F tools"}
    rows = []
    for c in ["D", "E", "F", "piD", "piE", "piF"]:
        if c not in piv:
            continue
        g = s[s.config == c]
        ref = "D" if not c.startswith("pi") else "piD"
        dd, se = paired(piv[c], piv[ref]) if c != ref else (np.nan, np.nan)
        rows.append([lab[c], len(g), f3(g.g_eval.mean()), f3(g.rouge1.mean()), f"{g.n_pred.mean():.1f}",
                     f"{g.n_rejected.mean():.1f}", "" if c == ref else f"{dd:+.3f} ({se:.3f})"])
        V[f"P3_{c}"] = f3(g.g_eval.mean())
        V[f"P3_{c}_D"], V[f"P3_{c}_SE"] = sg(dd), f3(se)
        V[f"P3_{c}_REJ"] = f"{g.n_rejected.mean():.1f}"
        V[f"P3_{c}_ZERO"] = int((g.n_pred == 0).sum())
    if "piD" in piv:
        dd, se = paired(piv["piD"], piv["D"])
        V["P3_PID_VS_D"], V["P3_PID_VS_D_SE"] = sg(dd), f3(se)
    return {"T_P3": table(["harness and tools", "tables", "insight score", "ROUGE-1", "insights", "rejected",
                           "vs D / vs pi D (se)"], rows,
                          "Table 7. InsightBench, 100 tables, analyst Qwen3-Coder-30B. Insight score: open judge "
                          "(GLM-4.5-Air), G-Eval-style best match of each planted insight, scaled to 0&ndash;1. "
                          "Top: our tool loop. Bottom: the same tools served to the pi coding agent.", "wide")}


# ------------------------------------------------------------------------------------------------- RelBench capability
SHORT = {"rel-avito/user-clicks": "avito clicks", "rel-avito/user-visits": "avito visits",
         "rel-event/user-ignore": "event ignore", "rel-event/user-repeat": "event repeat",
         "rel-f1/driver-dnf": "f1 dnf", "rel-f1/driver-top3": "f1 top3", "rel-hm/user-churn": "hm churn",
         "rel-trial/study-outcome": "trial outcome"}


def p5(runs: Path, V: dict) -> dict:
    ds = [d for d in sorted(runs.glob("*V4_relbench_capability")) if (d / "capability_rows.jsonl").exists()]
    d = max(ds, key=lambda x: len(jl(x / "capability_rows.jsonl")))
    r = pd.DataFrame(jl(d / "capability_rows.jsonl"))
    r["auroc"] = r.auroc.where(r.submitted)
    piv = r.groupby(["task", "config"]).auroc.mean().unstack()[["D", "E", "R", "ER"]]
    tasks = [t for t in SHORT if t in piv.index]
    rows = [[SHORT[t]] + [f3(piv.loc[t, c]) for c in piv] for t in tasks]
    mean = piv.loc[tasks].mean()
    rows.append(["<b>mean</b>"] + [f"<b>{f3(mean[c])}</b>" for c in piv])
    rows.append(["tool calls / episode"] + [f"{r[r.config == c].tools.map(lambda x: sum(int(v) for v in x.values())).mean():.1f}"
                                            for c in piv])
    rows.append(["seconds / episode"] + [f"{r[r.config == c].seconds.mean():.0f}" for c in piv])
    V.update(P5_T=len(tasks), P5_REP=int(r.repeat.max()) + 1, **{f"P5_{c}": f3(mean[c]) for c in piv})
    for c in ["E", "R", "ER"]:
        V[f"P5_{c}_WIN"] = f"{int((piv.loc[tasks, c] > piv.loc[tasks, 'D']).sum())} of {len(tasks)}"
    V["P5_SUB"] = f"{100 * r.submitted.mean():.0f}"
    return {"T_P5": table(["task", "D: LLM + SQL", "E: + Kumo Tabular-L", "R: + Kumo Relational", "ER: both"], rows,
                          f"Table 8. Prediction questions over RelBench databases (&ldquo;for each user, how likely is "
                          f"&hellip;&rdquo;). Official test AUROC, mean of {V['P5_REP']} episodes per cell; the agent "
                          f"writes features in SQL and calls the models as tools.")}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--runs", default="reports/runs")
    a = ap.parse_args()
    runs = Path(a.runs)
    V: dict = {}
    T = {}
    for f in (p0, p1, p2, p3, p5):
        T.update(f(runs, V))
    html = (HERE / "template.html").read_text()
    for k, v in {**T, **V}.items():
        html = html.replace("{{" + k + "}}", str(v))
    left = sorted(set(re.findall(r"\{\{(\w+)\}\}", html)))
    if left:
        print("unfilled:", left)
    (HERE / "paper.html").write_text(html)
    subprocess.run([CHROME, "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={HERE / 'vfd-agent-paper.pdf'}", f"file://{HERE / 'paper.html'}"],
                   check=True, capture_output=True)
    print("wrote", HERE / "vfd-agent-paper.pdf")


if __name__ == "__main__":
    main()
