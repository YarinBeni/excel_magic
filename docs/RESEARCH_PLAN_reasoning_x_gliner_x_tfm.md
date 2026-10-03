# Research plan: can reasoning LLMs, zero-shot text models (GLiNER/GLiClass) and tabular/relational FMs improve each other on real company data?

Draft, 3 October 2026. Setting: an automated BI analyst (LLM agent with SQL tools over company databases). Today it is
one reasoning LLM with tools. The question: what do the two other model families add, and in which direction should
the loop run?

## 1. What each family is good at, and where it stops

| family | examples (open) | strong at | weak at | cost |
|---|---|---|---|---|
| reasoning LLM | GLM-4.5/5, Qwen3, gpt-oss | planning, SQL, reading schema docs, explaining, judging text | numbers over many rows, consistency, hallucinated claims, long tables in context | high per token; seconds to minutes per step |
| zero-shot text model | GLiNER (NER, ~50 to 100+ labels per pass), GLiClass (classification, cross-encoder quality at ~10x speed) | turning free text into labels/entities at scale with no training; cheap verification of a text claim; reranking | no reasoning, no numbers, short inputs, domain words may need a small fine-tune | tiny (encoder, 100-500M); thousands of rows per second on one GPU |
| tabular / relational FM | TabPFN v2/2.5, TabICLv2, Kumo Tabular, Kumo Relational | prediction from rows with no training; feature importance by permutation; relational FM reads the DB graph | text columns (near chance on text-heavy rows: TEmBed J20), context limited to thousands of rows, no explanation | one GPU forward pass per batch; seconds |

The three families do not overlap. The LLM decides and explains; the text model makes text usable by the others and
checks text claims; the tabular model measures signal in the data. Our own results so far (papers 1 and 2) set the
frame: a frozen TFM is a strong, cheap "statistician" but cannot read text; an LLM-driven search on small tables is
limited by a noisy judge, and the fix is a better judge with more data, which a company database has.

## 2. The hypotheses

- **H1 (text -> table).** GLiNER/GLiClass labels of free-text columns, used as categorical features, raise a frozen TFM's
  accuracy more than dense text embeddings do, because TFMs use a few categorical columns better than hundreds of dense ones
  (our SVD/TabPFN result: dense factors through the TFM lose value).
- **H2 (table -> reasoning).** An LLM analyst with a "TFM probe" tool (fit in context, report AUROC, permutation
  importance, partial dependence per feature) produces insights that match the ground truth more often, and with fewer
  tokens, than an LLM that writes its own statistics in Python/SQL.
- **H3 (text -> reasoning).** A small zero-shot model as a verifier (does the text support the claim? which columns does
  the question mention?) cuts LLM errors on schema linking and claim checking at near-zero cost.
- **H4 (the loop).** Reasoning LLM proposes hypotheses as SQL feature tables; TFM judges them on held-out data; GLiNER
  structures text; accepted hypotheses become a reusable skill library per database. On relational benchmarks with
  ground truth, this beats the LLM alone and the TFM alone.

## 3. Benchmarks (all public, all have ground truth)

| benchmark | what it measures | used for |
|---|---|---|
| AutoGluon multimodal benchmark (Shi et al., NeurIPS 2021): 18 tabular datasets with 1 to 28 text columns, classification and regression | prediction with text + numbers | H1 |
| InsightBench (ServiceNow, ICLR 2025): 100 business datasets with planted insights, LLM judge, end-to-end analytics | insight quality of an analyst agent | H2, H4 |
| BIRD (dev, with evidence) and Spider 2.0-lite: text-to-SQL over real schemas | SQL correctness, schema linking | H3 |
| RelBench entity tasks (8 we already run) and rel-stack / rel-amazon (text-heavy) | relational prediction with official evaluator | H1 (text tables), H4 |
| TEmBed (IBM 2026) row tasks | text-heavy row embeddings | H1 control |

Caveat: BIRD and Spider 2.0 labels have known error rates (a 2026 audit reports 50 to 60% of annotations with issues on
the mini-dev / snow subsets), so use them for relative comparisons only. InsightBench's judge is an LLM; report
agreement with a second judge.

## 4. Experiments

### E0: tradeoff table (1 week)
Run each family alone on a slice of each benchmark. Record accuracy, seconds, tokens, GPU memory, and the failure
types. Output: Table 1 of the paper, and the choice of model sizes.

### E1: GLiNER/GLiClass as featurizer for a frozen TFM (H1, 1 to 2 weeks, cheap)
On the 18 multimodal datasets and 2 text-heavy RelBench tasks, compare TabPFN-2.5 / Kumo Tabular-L on:
(a) numeric + categorical only; (b) + MiniLM embeddings (384-d, and SVD-16 of them); (c) + GLiClass labels (the LLM
writes 5 to 20 label names per text column from the column name and 20 sample rows, GLiClass scores every row, top
label and its score become features); (d) + GLiNER entities (counts per entity type); (e) + an LLM's own per-row labels
on a 2,000-row subset (upper bound, cost logged). Metric: AUROC / RMSE, official splits. Baseline: AutoGluon's
multimodal result from the paper.
Expected: (c) > (b) on small datasets; (e) ~ (c) at 100x the cost.

### E2: TFM probe as a tool for the analyst agent (H2, 2 weeks)
InsightBench, 100 datasets. Agent A: the current BI agent (LLM + pandas/SQL). Agent B: A + a `tabfm_probe(table, target)`
tool that returns held-out AUROC/RMSE, permutation importance, partial dependence and the top interactions, computed by
a frozen TFM in seconds. Agent C: B + GLiClass tool for text columns. Same LLM (GLM or Qwen3, vLLM), same budget of
tool calls. Metric: InsightBench's insight score (its judge plus a second open judge), tokens and minutes per dataset.
Expected: B finds planted numeric trends more often and earlier; C adds the text-driven insights.

### E3: zero-shot verifier for the SQL agent (H3, 1 week)
BIRD dev. (a) LLM alone; (b) GLiNER schema linking: labels = column names (bi-encoder variant for 100+ labels), the
extracted spans become the candidate columns passed to the LLM; (c) GLiClass claim check: after the agent writes its
answer, GLiClass scores "the question asks about {label}" over the schema to flag missing joins. Metric: execution
accuracy, tokens. Expected: small gain, large token saving, on wide schemas.

### E4: the closed loop on relational data (H4, 3 to 4 weeks)
RelBench tasks with text (rel-stack, rel-amazon) plus our 8. The LLM proposes feature tables as SQL over the database;
the frozen TFM scores each on a time-split validation set (the judge; tens of thousands of rows, so less noise than
TabArena's hundreds); GLiClass turns text columns into labels on demand; accepted features go into a per-database
library that later tasks on the same database can reuse. Compare: LLM alone (LightGBM on its own features), TFM on
hand features (paper 2: 0.727), Kumo Relational graph layer (paper 2: 0.748), the loop. Also measure transfer: does the
library built on rel-avito user-visits help user-clicks? Metric: official test AUROC, plus a written insight judged
against the feature importances.

## 5. Self-improvement, defined narrowly
"Improve oneself" is operationalised as the skill library in E4: accepted SQL features, GLiClass label sets per text
column, and TFM importance reports, stored per database and offered to the agent on later tasks. The test is transfer
within a database and across databases of the same domain. No fine-tuning of the LLM.

## 6. Order and cost
E1 first (cheap, all open, reuses our registry and RelBench code; a result either way is publishable as a note).
Then E2 (needs the InsightBench harness and vLLM; ~2 GPU-days). E3 in parallel (CPU-light). E4 last, it depends on
E1's featurizer and E2's probe tool. Everything runs on the cluster through the git runner. Models: GLM-4.5-Air /
Qwen3-Coder-30B via vLLM (one H200), GLiNER-medium / GLiClass-modern-base (CPU or shared GPU), TabPFN-2.5 and Kumo
Tabular-L / Kumo Relational (one H200).

## 7. What would make it a paper
A clear tradeoff table (E0), one strong pairwise result (E1 or E2), and the loop (E4) with an ablation that removes
each family. The novelty is the division of labour, measured on public benchmarks with ground truth, not a new model.

## Sources
GLiNER: https://github.com/urchade/GLiNER ; GLiClass: https://github.com/Knowledgator/GLiClass , paper
https://www.arxiv.org/pdf/2508.07662 ; InsightBench: https://arxiv.org/abs/2407.06423 ,
https://github.com/ServiceNow/insight-bench ; multimodal benchmark: https://arxiv.org/abs/2111.02705 ,
https://github.com/sxjscience/automl_multimodal_benchmark ; Spider 2.0: https://arxiv.org/abs/2411.07763 ,
https://github.com/xlang-ai/Spider2 ; BIRD/Spider annotation audit: https://arxiv.org/abs/2601.08778
