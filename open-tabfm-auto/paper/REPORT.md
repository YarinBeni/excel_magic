# Open TabFM-Auto: Self-Evolving Pipelines Around Frozen Open-Weight Tabular Foundation Models, With Any Agent and Any LLM

*Working report, written in the structure of Fu et al. (2026), "TabFM-Auto: Self-Evolving Pipelines for Tabular
Foundation Models" (arXiv 2609.37989). Numbers marked **[auto]** are regenerated from run directories by
`paper/fill_results.py` (see `paper/results_auto.md`); cells marked **TBD** are runs that are queued or not yet done.
Last regenerated: see the header of `results_auto.md`.*

## Abstract

Tabular foundation models (TFMs) predict strongly in a single forward pass but ignore column semantics. TabFM-Auto
showed that a coding agent that evolves the data pipeline around a *frozen* TFM, judged by cross-validation,
reaches the top of TabArena. The original work depends on a closed 400M-parameter model (TabFM) and on frontier
LLM agents (Claude Opus 5, Gemini 3.8 Flash) at a cost of ~$17.6K per sweep. We reimplement the method with
open components only and ask four questions the paper leaves open: (1) does the pipeline-search gain transfer to
open-weight TFMs of very different strength (TabPFN v2, TabPFN-2.5, TabICLv2, EXAONE-Tabular, Kumo Tabular S/M/L)?
(2) how much of the gain survives when the frontier LLM is replaced by an open 20–32B model served locally
(Qwen3-Coder, Qwen3, GLM-4.5-Air, gpt-oss), or by *no* LLM at all (a greedy search over generic operations)?
(3) does the agent harness matter (Claude Code vs. a minimal OpenAI-compatible tool loop vs. free CLI agents: pi,
Qwen Code, aider)? (4) can the same loop start from a relational database instead of a flat table? We release
`open-tabfm-auto` (MIT) with the paper's pipeline contract, a sandboxed judge, the TabArena protocol with the paper's
per-dataset numbers for direct comparison, and pluggable backbones, LLMs and harnesses. As a side result we study
whether the hidden states of frozen TFMs and relational FMs are useful graph-aware entity embeddings for retrieval.

## 1. Introduction

The paper's claim is a mechanism: keep the predictor frozen, let the LLM turn metadata (column names, task text) into
*code* around it, and judge every candidate with the same cheap, noise-free CV signal. Two things make the result
hard to use outside Google: the frozen model is not released, and the search spends thousands of frontier-LLM calls
per dataset. Our contribution is to separate the mechanism from those two ingredients and measure what each one is
worth. Concretely we vary, one axis at a time, the frozen backbone, the LLM, the agent harness and the data source,
while keeping the pipeline contract, the judge and the evaluation protocol identical to the paper.

**Contributions.** (i) An open implementation of the full TabFM-Auto loop. (ii) A backbone study across seven
open TFMs spanning ~1400–1960 TabArena Elo. (iii) An LLM/harness study including open models and a no-LLM baseline
that isolates the value of domain knowledge. (iv) A SQL mode where the agent also builds the training table.
(v) A retrieval study of frozen hidden states as entity embeddings, evaluated on a synthetic relational DB and on
RelBench `rel-hm` with the official evaluator.

## 2. Related work

Tabular foundation models and in-context learning (TabPFN, TabICL, TabFM, Kumo Tabular, EXAONE-Tabular); AutoML and
LLM feature engineering (CAAFE, FeatLLM, OCTree); MLE agents (AIDE, R&D-Agent, MLEvolve); relational foundation
models (KumoRFM-2 and NVIDIA's KumoRelational, Relational Transformer, OpenRFM). The paper's Section 2 covers the
first three; we add the relational line because our SQL mode and embedding study touch it.

## 3. Method

### 3.1 Problem formulation (unchanged)
Given D_train, X_test and metadata M, find pipeline P = (Φ_clean, Φ_feat, S_ctx, Ψ_post) maximising the 3-fold CV metric
of the frozen model f_θ* on the training split (paper Eq. 1). Test rows are never visible during search.

### 3.2 Pipeline contract (unchanged)
`pipeline.py` defines `preprocess`, `engineer` (≤ 500 columns), `sample` (context views, each fitted separately and
averaged), `postprocess`, and `MODEL_KWARGS`. The identity pipeline P0 starts every search (paper Fig. 8).

### 3.3 Judge
`tabfm-eval`: 3-fold CV on the training split in a subprocess with network variables stripped and a timeout;
every call appends a record (score, folds, features, timing, traceback) and snapshots the candidate. The best-CV
candidate is frozen after the budget and scored once on held-out data. Isolation is best-effort (no bubblewrap).

### 3.4 What we vary (the expansion)
| axis | paper | here |
|---|---|---|
| frozen backbone | TabFM 400M (closed) | TabPFN v2 (11M), TabPFN-2.5, TabICLv2 (28M), EXAONE-Tabular (21M), Kumo Tabular S (28M) / M / L (215M) |
| LLM | Claude Opus 5, Gemini 3.8 Flash | Claude Sonnet; Qwen3-Coder-30B-A3B, Qwen3-32B, GLM-4.5-Air, gpt-oss-20b (vLLM, local); **none** |
| harness | Claude Code, Codex, Antigravity | Claude Code; a 5-tool OpenAI-compatible loop; pi, Qwen Code, aider (CLI); greedy heuristic (no LLM) |
| data source | flat table (+ aux files) | flat table; **SQL database** (agent writes `build_table.sql` with a temporal cutoff, then searches) |
| budget | 96 evals / 6 h / H100 | 16–24 evals / ≤ 90 min per dataset (CPU or one H200) |

The no-LLM harness (`heuristic`) is a greedy coordinate search over the paper's "generic" operations (sentinel-to-NaN,
winsorizing, count encoding of ID-like columns, log of skewed features, crosses of the top-k informative features,
truncated SVD, pruning, balanced / multiple context views, prior correction, temperature, log1p target, more
ensemble members). It is the control for "how much of the gain is domain reading".

## 4. Experimental setup

**Datasets.** (a) Bundled and synthetic tables with planted structure (a physics-style regression whose target depends
on a Strouhal-type ratio; an entity-ID classification whose signal is pairwise co-occurrence), used for cheap,
interpretable comparisons. (b) TabArena v0.1 through the paper's protocol: search on fold 0's training split, score
P0 and P* on all official splits, compare per dataset with the paper's Table 9 (TabFM, TabFM-Auto). First wave:
the 17 datasets with ≤ 2,500 rows (30 official splits each).
**Metrics.** TabArena's: 1−AUROC (binary), log loss (multiclass), RMSE (regression); Bradley-Terry Elo when a pool
is available (`tabfm_auto.benchmarks.elo`).
**Compute.** Development on a 4-core CPU container; cluster runs on H200 nodes via a git-driven Slurm runner.

## 5. Results

### 5.1 Frozen backbones with the identity pipeline (T1 **[auto]**)
On five small tables (H200, 3-fold CV) Kumo Tabular-L is best on 3/5 and TabICLv2 on 2/5; TabPFN v2, the only model
with a non-HF mirror, is last among the TFMs on 4/5; all TFMs beat HistGB by 1.3–4× (full table in `docs/RESULTS.md`).
Inference cost is not a constraint at this scale: 0.4–3.7 s per 3-fold CV for every model.

### 5.2 Pipeline search: does the gain survive open LLMs and no LLM? (T2 **[auto]**)
Local, TabPFN v2 as backbone, held-out test:

| dataset | P0 | Claude Sonnet (Claude Code) | no-LLM heuristic | HistGB |
|---|---|---|---|---|
| synth_physics (RMSE) | 0.1108 | **0.0830** (−25.1%, 4 evals, $0.15) | 0.0852 (−23.1%, 20 evals) | 0.3277 |
| synth_entities (1−AUROC) | 0.4676 | TBD | 0.4266 (−8.8%) | 0.4585 |
| breast_cancer (1−AUROC) | 0.0061 | TBD | 0.0061 (0%) | 0.0108 |

The Sonnet agent rediscovered the planted physical ratio in four evaluations; the heuristic recovers most of that
gain because pairwise ratios of the top features are in its library. Open LLMs (J3, J5) and CLI harnesses (J6): **TBD**.

### 5.3 TabArena protocol vs the paper (T3 **[auto]**)
Wave 1 (the 17 datasets with <= 2,500 rows, all 30 official splits each, frozen Kumo Tabular-S, **no LLM**):
mean relative test-error gain **+0.5%** against **+4.6%** for the paper's Opus 5 agent on the same datasets; P* beats P0
on 9/17. The open backbone's identity pipeline is within a few percent of the paper's unreleased TabFM on 16/17
datasets (better on 5). The generic search recovers a third to a half of the paper's gain on datasets with generic
structure (anneal +8.5% vs +17.6%; airfoil +6.5% vs +14.6%) and loses on small noisy tables where fold-0 CV gains do
not transfer (diabetes -1.7%, maternal_health_risk -2.7%), the regime where the paper also reports its own losses.
Full table: `docs/RESULTS.md`. Open-LLM and CLI-agent runs on the same protocol: **TBD (J3/J5/J6)**.

### 5.4 Backbone transfer (T4 **[auto]**)
The pipeline found for TabPFN also helps HistGB on the physics table (0.328 → 0.181, −45%), the paper's Section B.3
effect. Across the open TFMs the picture differs from the paper: pipelines found around the *weakest* backbone
(TabPFN v2) transfer unevenly to stronger ones (physics: +15.5% on TabPFN, −11.3% on TabPFN-2.5, ≈0 on Kumo;
entities: +26% on Kumo-S, −18% on Kumo-M). The paper transfers from the strongest model downwards; we transfer upwards,
and features that a weak model needs are often redundant or harmful for a stronger one. Search should run per backbone.

### 5.5 SQL mode
On a synthetic shop database (customers / products / orders / order items / tickets; churn in the 90 days after a
cutoff), the agent's SQL-built table (38 features) reached held-out 1−AUROC 0.293 with P0 and 0.275 after search,
beating a hand-written 9-feature table (0.317) and HistGB on the agent table (0.314). Total LLM cost ≈ $0.30.

### 5.6 Extension: frozen hidden states as entity embeddings (T5, T6 **[auto]**)
Synthetic shop DB, 3 seeds, latent-segment retrieval P@10 (chance 0.167): row-only 0.161; hand aggregates 0.726;
frozen TabPFN over the aggregates with a single random in-context target 0.49 ± 0.34 (unstable); the same with
k-means pseudo-labels as target **0.824 ± 0.021**; a GNN trained on the DB 0.877 ± 0.007; OpenRFM hidden state 0.249
(random weights 0.168). The in-context target decides what the embedding keeps: a downstream label collapses it onto
that label (0.18). RelBench `rel-hm` user-item-purchase with the official evaluator (J7): **TBD**; trivial baselines
reproduced on val (GlobalPopularity 0.34, PastVisit 1.90, MAP@12 × 100).

## 6. Ablations and analysis (planned)
LLM vs no-LLM gap per dataset category (domain-readable vs anonymised schemas, paper Table 6); budget curves
(paper Fig. 4); harness effect at fixed LLM; backbone strength vs relative gain; cost per dataset.

## 7. Limitations
Budgets are 4–6× smaller than the paper's; isolation is best-effort; TabPFN v2 is far below the paper's backbone;
open-LLM tool calling quality varies by model and vLLM parser; the synthetic tables are easy to over-interpret and are
there for diagnosis, not as benchmarks.

## Reproducibility
`open-tabfm-auto` (MIT): `tabfm-auto search|db|models`, `examples/05_tabarena_protocol.py`, `scripts/slurm_tabarena.sh`;
every run writes `metrics.json`, the agent stream and every candidate pipeline. `paper/fill_results.py` rebuilds all
tables from those files.
