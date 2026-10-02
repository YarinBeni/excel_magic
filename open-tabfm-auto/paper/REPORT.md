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

**Findings so far (2026-10-02, cluster).** With a frozen open backbone (Kumo Tabular-S) the identity pipeline is already
within a few percent of the paper's TabFM on 16/17 small TabArena datasets. The no-LLM greedy search recovers a mean
+0.5% on those 17 (paper: +4.6% with Opus 5). Open LLMs served locally differ far more by *harness* than by model: on
the entity-aggregation task the same Qwen3-Coder-30B gives +50% through pi, +42% through aider, +26% through Qwen Code
and ~0% through a minimal tool loop; GLM-4.5-Air-FP8 reaches +44% even in the minimal loop. Every open setup overfits
3-fold CV on the smallest table. On retrieval, frozen TFM hidden states are usable segment-level entity embeddings on a
synthetic DB (0.82 P@10 vs GNN 0.88) when the in-context target is chosen well, but on RelBench rel-hm at item level
they lose to their own input (0.42 vs 0.83 MAP@12×100 for the raw interaction factors) and to sparse cosine on the
purchase matrix (1.20).

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
gain because pairwise ratios of the top features are in its library.

**Open LLM, first results (cluster, Kumo Tabular-S frozen, our OpenAI-compatible tool loop over vLLM 0.16, Qwen3-Coder-30B-A3B,
two independent runs J3 / J5, held-out test, gain = relative error reduction):**

| dataset | P0 (J3 / J5) | Qwen3-Coder gain (J3 / J5) | evals | no-LLM heuristic (J1, same backbone) |
|---|---|---|---|---|
| synth_physics (RMSE) | 0.0869 / 0.0865 | +1.3% / +2.9% | 8 / 12 | +0.6% / 0.0% |
| synth_entities (1−AUROC) | 0.4387 / 0.4545 | −1.1% / +1.0% | 9 / 12 | +29.0% / +24.6% |
| breast_cancer (1−AUROC) | 0.0051 / 0.0044 | −2.9% / −56.7% | 16 / 9 | 0% (TabPFN) |
| TabArena-Lite airfoil / blood / credit-g | | −0.9% +2.5% −0.3% / +2.3% −0.9% −0.4% | 24 | |

Qwen3-Coder uses the tools correctly (describe, write, eval, finish; ~100k input tokens, ~7k output per search, 1.5 min
on one H200) but its pipelines barely move the frozen Kumo-S, and on breast_cancer (569 rows) the CV-best candidate loses
badly on test. On synth_entities the LLM misses the entity-aggregation trick the greedy heuristic finds. One J3 run
looped `run_eval` 38 times without changing the pipeline; the harness now refuses an unchanged re-evaluation. gpt-oss-20b
crashed the loop by reading a parquet file (tool errors now return to the model); it, Qwen3-32B, GLM-4.5-Air-FP8 and the
CLI harnesses (pi, Qwen Code, aider) are queued/running.

**CLI harness, first result (J6, aider + Qwen3-Coder-30B-A3B over vLLM, same frozen Kumo-S, 8 write→eval rounds):**
synth_physics −2.2%, breast_cancer −40.0%, and synth_entities **+41.7%** (1−AUROC 0.384 → 0.224, HistGB 0.473): the
single-turn aider loop wrote the entity frequency encodings (role×resource pair counts, manager request counts, their
ratios) that the task's description hints at, plus a 4-view class-balanced context sample. Same LLM, different harness:
the tool-loop runs above never found this. CV trajectory 0.470 → 0.317 on the second eval, then flat.

**Qwen Code (same LLM, J6 task 1, 128k server context):** synth_physics +1.0%, synth_entities **+25.5%** (0.459 → 0.342),
breast_cancer −16.1%; 22 API calls, 616k prompt tokens, 10 shell evals, ~140 s per search. Ranking on the entity task,
all with Qwen3-Coder-30B: aider loop +41.7% > Qwen Code +25.5% ≈ no-LLM heuristic +25–29% > our tool loop ±1%.

**pi (J6 task 0, same LLM):** synth_physics −0.9%, synth_entities **+50.3%** (0.452 → 0.224, 6 evals, 89 s), breast_cancer
−17.1%. Final entity-task ranking with one fixed open LLM (Qwen3-Coder-30B) and one frozen backbone: pi +50.3% > aider
+41.7% > Qwen Code +25.5% ≈ heuristic +25–29% > our tool loop ±1%; all three CLI agents converge on the same
role×resource frequency encodings, and all lose on breast_cancer.

**Summary of the LLM/harness study (synthetic tasks, frozen Kumo Tabular-S, held-out test, relative error reduction):**

| setup | physics | entities | breast_cancer | wall / search |
|---|---|---|---|---|
| no LLM, greedy heuristic (J1) | +0.6 / 0.0% | +29.0 / +24.6% | 0% | ~10 min |
| tool loop + Qwen3-Coder-30B (J3, J5) | +1.3 / +2.9% | −1.1 / +1.0% | −2.9 / −56.7% | 1.5 min |
| tool loop + Qwen3-32B | −0.2% | −3.8% | −2.9% | 17–30 min |
| tool loop + gpt-oss-20b | +4.1% | −8.9% | −7.9% | ~5 min |
| tool loop + GLM-4.5-Air-FP8 | **+5.1%** | +43.6% | −38.9% | 5 min |
| Qwen Code + Qwen3-Coder-30B | +1.0% | +25.5% | −16.1% | 2.3 min |
| aider loop + Qwen3-Coder-30B | −2.2% | +41.7% | −40.0% | 2.5–4 min |
| pi + Qwen3-Coder-30B | −0.9% | **+50.3%** | −17.1% | 1.5–2.5 min |
| Claude Code + Sonnet (local, TabPFN v2) | +25.1% | TBD | TBD | $0.15 |

Two regularities. The *harness* decides whether the entity encodings are found: three different CLI agents converge on
the same role×resource frequency features with the same 30B model that finds nothing in a minimal tool loop; the
larger GLM finds them in the minimal loop. And every open setup loses on the 569-row breast_cancer table, where the
3-fold CV signal is too noisy for a 16-eval search; the paper's own small losses (credit-g) are the same regime.

**gpt-oss-20b (J5 task 3, our tool loop):** synth_physics −1.3% (2 evals), synth_entities −9.7%, breast_cancer +7.9%;
vLLM's harmony parser leaked channel markers into tool names (`run_eval<|channel|>commentary`), wasting calls; the
harness now normalises names. Its TabArena-Lite run died on an empty vLLM response (now guarded); rerun queued.

**gpt-oss-20b rerun with the fixed harness (job 50434):** synth_physics **+4.1%** (0.0869 → 0.0834), synth_entities −8.9%,
breast_cancer −7.9%; TabArena-Lite airfoil +2.0%, blood −1.1%, credit-g −0.2%. It writes far more candidates than it
evaluates (33 writes / 8 evals on entities) and still misses the aggregation features.

**GLM-4.5-Air-FP8 (J5 task 2, our tool loop, 106B-A12B, one H200):** synth_physics **+5.1%**, synth_entities **+43.6%**
(0.397 → 0.224, the best entity result of any run, matching aider's), breast_cancer −38.9%; TabArena-Lite airfoil −0.9%,
blood +1.4%, credit-g +1.4%. It uses the full 16-eval budget (write → eval every turn, ~500k input / 25k output tokens,
~5 min per search). Among open LLMs in our tool loop the order on the synthetic tasks is GLM-4.5-Air > gpt-oss-20b >
Qwen3-Coder-30B; the small-data overfit on breast_cancer is common to all of them.

**Qwen3-32B (J5 task 1, reasoning model, hermes parser):** synth_physics −0.2% (it answered in text after one
`describe_data` call and never evaluated), synth_entities −3.8% (37 writes, 16 evals, 30 min, 830k input / 88k output
tokens), breast_cancer −2.9%; TabArena-Lite airfoil +0.7%, blood +0.6%, credit-g +0.3%. Weakest of the sweep.

### 5.3 TabArena protocol vs the paper (T3 **[auto]**)
Wave 1 (the 17 datasets with <= 2,500 rows, all 30 official splits each, frozen Kumo Tabular-S, **no LLM**):
mean relative test-error gain **+0.5%** against **+4.6%** for the paper's Opus 5 agent on the same datasets; P* beats P0
on 9/17. The open backbone's identity pipeline is within a few percent of the paper's unreleased TabFM on 16/17
datasets (better on 5). The generic search recovers a third to a half of the paper's gain on datasets with generic
structure (anneal +8.5% vs +17.6%; airfoil +6.5% vs +14.6%) and loses on small noisy tables where fold-0 CV gains do
not transfer (diabetes -1.7%, maternal_health_risk -2.7%), the regime where the paper also reports its own losses.
Full table: `docs/RESULTS.md`. Open-LLM and CLI-agent runs on the same protocol: **TBD (J3/J5/J6)**.


**With LLM agents (J8, same protocol and backbone, 16 evals / 40 min per dataset):** pi + Qwen3-Coder-30B finished all
17 datasets with mean gain −0.9% (median 0.0%, P* beats P0 on 8/17, better than the heuristic on 7/17); the GLM-4.5-Air
tool loop is at 6/17 (mean −2.9% so far, dragged by maternal_health_risk −13.8%). Neither matches the heuristic's +0.5%,
and all three are an order of magnitude below the paper's +4.6%. The paper's big wins are exactly where the free LLM
agents lose: anneal (paper +17.6%, heuristic +8.5%, pi −5.2%) and Marketing_Campaign (paper +15.8%, pi −14.2%); the
LLM picks a CV-best candidate that does not hold on the other 29 splits. Per dataset:

| dataset | metric | heuristic gain | pi+Qwen3-Coder gain | GLM loop gain | paper Opus 5 gain |
|---|---|---|---|---|---|
| anneal | logloss | +8.5% | -5.2% | -2.7% | +17.6% |
| Marketing_Campaign | 1-auroc | +0.1% | -14.2% |  | +15.8% |
| airfoil_self_noise | rmse | +6.5% | +3.7% |  | +14.6% |
| hazelnut-spread-contaminant-detection | 1-auroc | +1.3% | +0.2% |  | +8.7% |
| diabetes | 1-auroc | -1.7% | +0.8% | +1.3% | +6.9% |
| concrete_compressive_strength | rmse | -1.3% | +0.5% |  | +3.8% |
| MIC | logloss | +1.5% | -0.1% |  | +2.4% |
| website_phishing | logloss | +0.4% | -0.1% |  | +1.8% |
| healthcare_insurance_expenses | rmse | -1.3% | +0.8% |  | +1.7% |
| maternal_health_risk | logloss | -2.7% | +0.2% | -13.8% | +1.6% |
| Another-Dataset-on-used-Fiat-500 | rmse | -0.7% | -0.4% |  | +1.4% |
| Is-this-a-good-customer | 1-auroc | -2.3% | -0.0% |  | +1.4% |
| QSAR_fish_toxicity | rmse | +0.2% | -0.1% | +0.0% | +0.6% |
| blood-transfusion-service-center | 1-auroc | +0.2% | -0.1% | -0.4% | +0.4% |
| credit-g | 1-auroc | -0.2% | -2.2% | -1.7% | +0.2% |
| Fitness_Club | 1-auroc | +0.7% | +0.2% |  | +0.1% |
| qsar-biodeg | 1-auroc | -1.2% | +0.1% |  | -0.3% |

Reading: at this budget the open LLM agents do not beat a generic greedy search on TabArena. Their advantage shows
only where the task text names an entity structure (§5.2). J10 (64 evals + repeated CV on the five high-gain datasets)
and J11 (repeated-CV judge on the small tables) test whether the gap is budget and judge noise rather than the model.

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
(random weights 0.168); Kumo Relational (frozen, cluster) 0.62 ± 0.05 with k-means target vs 0.33 ± 0.13 random target.
The in-context target decides what the embedding keeps: a downstream label collapses it onto that label (0.18).

RelBench `rel-hm` user-item-purchase, official evaluator, MAP@12 × 100 (J7, job 50266, 74,575 val / 67,144 test
queries, 365-day history, k = 50): GlobalPopularity 0.34 / 0.29 (val / test), PastVisit 1.90 / 2.20 (published test
rows: GlobalPop 0.30, PastVisit 0.89, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93). Training-free user-kNN
over our customer embeddings: row-only 0.14, hand aggregates 0.26, TabPFN(agg, k-means target) 0.23, TabPFN(agg,
random target) 0.17, TabPFN(row, k-means) 0.14 on test; every PastVisit+kNN hybrid 2.19, i.e. the fill never helps.
**Negative result, attributed** (J7c, job 50341): with the same kNN scoring, user-kNN on the raw purchase matrix
reaches 1.20 test (4× popularity, above the published LightGBM 0.38 and GraphSAGE 0.80), item-kNN 1.07, and
PastVisit+kNN-CF[purchase matrix] 2.22 is the only hybrid above PastVisit. So the retrieval step is fine; the customer-level
features, with or without the frozen TFM on top, do not carry item-level co-purchase structure. J7d feeds the interaction
signal itself (SVD factors of the purchase matrix, plus aggregates) through the frozen TabPFN and compares it with the raw
factors; that is the last version of the hypothesis still open at item level.

**J7d closes it** (job 50399): with 64 SVD factors of the purchase matrix as input, plain kNN scores 0.83 test; the frozen
TabPFN hidden state over the same input scores 0.42 (k-means target) / 0.48 (random target), and adding the customer
aggregates to the factors already drops them to 0.65. At item level on a real purchase graph the frozen TFM embedding is
worse than its own input, which is worse than sparse cosine on the interaction matrix (1.20). The hypothesis survives
only at segment level (synthetic, with a well-chosen in-context target), which is how the paper plan now frames it.

**Entity-level probe on real data (J12, rel-hm user-churn, 20k customers at one timestamp, 5-fold kNN probe, AUROC):**
supervised HGB on the aggregates 0.673; kNN on the raw aggregates 0.653; kNN on the frozen TabPFN hidden state over the
same aggregates 0.648 (k-means target) / 0.643 (random); SVD factors 0.590; row-only 0.520. The same frozen TabPFN used as a classifier with the labels in context reaches 0.672, equal to HGB: the pseudo-target costs 2.4 AUROC points, which is the whole gap between "frozen embedding" and "frozen model". The frozen embedding
equals its input and does not add what a supervised model extracts. Together with J7c/J7d the retrieval claim reduces
to: a frozen TFM with a k-means in-context target is a target-agnostic entity representation that preserves the structure
of the features it is given (and recovers planted segments from them), at zero training cost; it is not a better
representation than those features, and it is not an item-level retrieval embedding.

## 6. Ablations and analysis
- **Harness at fixed LLM** (§5.2): pi +50% > aider +42% > Qwen Code +26% > tool loop ~0% on synth_entities with
  Qwen3-Coder-30B. The CLI agents read the task text and the data files themselves and iterate on eval output; the
  minimal loop exposes the same information through tools, yet the 30B model does not use it there.
- **Affordances vs agency (J13)**: giving the minimal tool loop what the CLI agents see (file listing, pipeline.py,
  the data summary, in the first message) does not close the gap: Qwen3-Coder-30B then scores −0.7% / −4.9% / −23.5%
  on physics / entities / breast_cancer (plain loop: +1.3 to +2.9% / −1.1 to +1.0% / −2.9 to −56.7%; pi: −0.9% / +50.3% /
  −17.1%). TabArena-Lite airfoil +3.3%, blood +0.4%, credit-g +1.4%. The harness effect is in how the agent iterates
  (reading eval output, editing in place, shell access), not in the information it starts with.
- **Judge noise (J11)**: a repeated 3-fold judge (3 repeats = 9 folds) at the same budget on the five TabArena tables
  with < 1000 training rows moves the heuristic's mean gain from +1.4% to +2.1% (anneal +8.5% → +10.5%, diabetes
  −1.7% → +1.2%) and pi + Qwen3-Coder from −1.4% to −0.5% on the four done so far; on the 569-row breast_cancer table pi's
  loss shrinks from −17.1% to −5.4% and the heuristic's stays at 0. Repeated CV is cheap (the folds are ~1 s each on an
  H200) and is now the recommended default for n < 1000 (`--cv-repeats auto`).
- **Budget (J10, paper Fig. 4)**: raising the budget from 16–24 to 64 evaluations (and `--cv-repeats auto`) on the five
  datasets where the paper gains most (anneal, Marketing_Campaign, airfoil, hazelnut, diabetes; paper mean +12.7%) does
  not help: heuristic +2.9% → +1.1%, pi + Qwen3-Coder −2.9% → −4.3%, GLM-4.5-Air loop +0.1% → −8.4% (Marketing_Campaign
  −47.7%). The LLM agents also stop long before the budget (8–22 evaluations); only the heuristic uses 31–44.
- **Where the gap comes from (split transfer)**: scoring the same pipelines per official split shows that a candidate
  chosen by CV on the search split (r0f0) is often *not* better on the other 29 splits of the same data. Mean gain on
  r0f0's own test vs the other 29: heuristic −0.5% / −0.1%, pi +1.4% / −1.3%, pi at budget 64 −0.1% / −6.4%, GLM −1.3% /
  −1.2%; per dataset, anneal: heuristic +11.5% on r0f0 but −0.5% elsewhere, GLM +13.2% / −5.5%; Marketing_Campaign: pi
  +4.3% / −24.8%. The fraction of splits a candidate improves is 0.39–0.57, i.e. a coin flip. With 500–2500-row tables a
  3-fold CV difference of ±1–2% is inside the noise, so the search mostly selects noise, and a bigger budget selects more
  of it. Repeated CV (J11) is the only lever we found that moves this, and only partly.
- **LLM at fixed harness** (§5.2): GLM-4.5-Air (106B-A12B) > gpt-oss-20b > Qwen3-Coder-30B-A3B > Qwen3-32B in the
  tool loop; the reasoning model spends its budget on rewrites (37 writes / 16 evals) and finds nothing.
- **LLM vs no LLM**: the greedy heuristic matches the mid-tier CLI agents on the entity task (+25–29%) because
  group-by frequency encodings are in its operation library; it cannot discover the physics ratio that domain text
  suggests, which Sonnet found in four evals. Open LLMs found neither on physics (≤ +5%).
- **Backbone strength vs gain** (§5.4): pipelines found for TabPFN v2 transfer upward unevenly (Kumo-M −18% on
  entities, Kumo-L +5%), supporting the paper's conclusion that search should be run per backbone.
- **Small-data overfitting**: on breast_cancer every LLM setup's CV-best candidate is worse on test (−3% to −57%)
  with a single 3-fold judge; the repeated judge above removes most of it.
- **Cost**: one H200 serves a 30B-A3B coder at ~1.5 min per 16-eval search and GLM-4.5-Air-FP8 at ~5 min; the
  17-dataset TabArena wave costs ~1–2 GPU-hours per setup versus the paper's $17.6K sweep (J8, running).

## 7. Limitations
Budgets are 4–6× smaller than the paper's (16–24 evals vs ~100); isolation is best-effort; the frozen backbone (Kumo
Tabular-S, ~1850 Elo) is below the paper's TabFM on most datasets; open-LLM tool calling depends on the vLLM parser
(harmony leaked channel markers into tool names, Qwen Code needs a 128k server context); the synthetic tables are
diagnostic, not benchmarks; the retrieval study covers one real benchmark (rel-hm) and one TFM (TabPFN v2) at item
level, so the negative result is about that setting, not about all relational FMs.

## Reproducibility
`open-tabfm-auto` (MIT): `tabfm-auto search|db|models`, `examples/05_tabarena_protocol.py`, `scripts/slurm_tabarena.sh`;
every run writes `metrics.json`, the agent stream and every candidate pipeline. `paper/fill_results.py` rebuilds all
tables from those files.
