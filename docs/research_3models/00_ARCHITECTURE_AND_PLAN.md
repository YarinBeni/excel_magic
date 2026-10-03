# Verified, Fast, Deep: a three-model harness for data-analysis agents. Architecture and experiment plan

Draft 1, 3 October 2026. Built on the four research reports in this folder (01 GLiNER/GLiClass, 02 tabular and
relational FMs, 03 LLM analyst agents and benchmarks, 04 small verifiers). Supersedes the first sketch in
`docs/RESEARCH_PLAN_reasoning_x_gliner_x_tfm.md`.

## 1. What the evidence says (the design constraints)

| finding | source | design consequence |
|---|---|---|
| 81% of wrong SQL is a schema or meaning error, not syntax; enterprise schemas have 700-3,000 columns | 03 | schema linking is the first job of the small models |
| "query executes" predicts correctness at chance (AUROC 0.50); logprob / self-consistency 0.67; LLM judges 0.72-0.78; two judges 0.82 | 04 | execution is not verification; a cascade of independent checks is needed |
| verbalised agent confidence: 22% actual vs 77% claimed success | 04 | never use the LLM's own confidence as the signal |
| small verifiers match large ones only when trained in-domain (MiniCheck 770M ~ GPT-4; Weaver 400M keeps 98.7%); zero-shot step checkers collapse (ProcessBench 31.5 F1) | 04 | GLiClass = triage and routing zero-shot; real step judging only after fine-tuning on our own traces |
| GLiClass: 20-200 ms, 1 -> 128 labels costs 8% speed; scores uncalibrated (order and wording sensitive, saturation); no "none" class; 512-token window | 01 | many-label triage per step; per-label calibration (100-300 examples); one hypothesis per call for entailment; chunk |
| no published numbers for GLiClass / NLI as SQL or step verifiers | 01, 03 | open gap: a measured contribution in itself |
| long-horizon analysis decays ~47 points from early to late turns; state errors are 52-69% of failures | 03 | explicit state (a ledger of verified facts) instead of a growing chat |
| semantic layer: 84-90% -> 98-100%; multi-candidate SQL + execution checks +6-20 pts at 5-30x calls | 03 | a curated metric layer per database, built and verified once, reused (the skill library) |
| TFMs win small-to-medium IID data, lose temporal / large / text-rich slices to GBDT (BeyondArena); CPU unusable | 02 | the deep tool always reports a GBDT check beside the TFM on temporal or >100K-row tasks; runs on a GPU worker |
| TabPFN-extensions give SHAP, PD, CRT p-values, outlier scores, full predictive distributions; Kumo Relational predicts over a DB in ~1 min | 02 | the deep tool returns statistics, not rows: the LLM sees conclusions about millions of rows in a few hundred tokens |
| benchmarks carry heavy label error (BIRD Mini-Dev 52.8%, Spider 2.0-Snow 62.8%) | 03 | evaluate on corrected subsets where they exist; report both; never tune on test |
| commercial-safe weights: TabPFN v2, TabICLv2, Mitra, Kumo Tabular (OpenMDW); GLiNER v2.x / GLiClass Apache-2.0 | 01, 02 | the company deployment path uses only these |

## 2. The architecture

```
 question
    |
    v
 [G1 intent + schema link]  GLiClass: answerable / ambiguous / out of scope ; GLiNER bi-encoder: question spans ->
    |                       candidate columns and values (labels = column descriptions from the metric layer)
    v
 [L1 plan]  LLM writes a short plan: steps typed as SQL / DEEP / CHECK / ANSWER, each with its expected output
    |
    v  for each step
 [G2 step triage]  GLiClass multi-label over an error taxonomy: unseen column, missing time filter, join fan-out risk,
    |              repeats a step, aggregate without grain, irreversible, off-task ...  calibrated scores
    |   high risk -> L1 revises the step ; middle band -> J (LLM judge) ; low -> run
    v
 [X execute]  SQL on the database (read-only) | DEEP: TFM tool (below) | text columns -> G4 featurizer
    |
    v
 [L2 interpret]  LLM writes claims about the result, one fact per line
    |
    v
 [V verify each claim]  numeric claim -> recomputed by SQL, must match exactly (deterministic)
    |                   text claim   -> GLiClass entailment (claim as hypothesis, result as text) ; MiniCheck as the
    |                                   trained alternative ; uncertain band -> J
    v
 [S state ledger]  only verified facts enter the ledger; the LLM plans the next step from the ledger, not from chat
    |
    v
 [ANSWER]  report = ledger facts + confidence per fact (calibrated) + the SQL behind each number
```

**The DEEP tool (tabular / relational FM).** Inputs: a task table defined by SQL with a time cutoff, a target, optional
text columns. It samples the context (stratified, time-aware), runs a frozen TFM with cached context (TabICLv2 or Kumo
Tabular for commercial use; TabPFN-2.5/3 for research), under a time split, and returns JSON: held-out metric, top
features with SHAP and CRT p-values, partial-dependence summaries, top interactions, outlier rows (ids only), drift
score between periods, and the same metric from LightGBM when the task is temporal or larger than 100K rows. For
multi-table questions it runs Kumo Relational over the database and returns its graph-layer probe (paper 2: the graph
layer is the better embedding). The LLM never sees raw rows beyond samples.

**G4 featurizer.** For a text column the LLM proposes 5-20 label names from the column name and 20 sample values;
GLiClass scores every row in batches; the per-label scores become numeric columns for the DEEP tool.

**J (escalation judge).** A second open LLM from a different family, called only for the uncertain band, so its cost
is bounded.

**Learning across tasks (the "improve itself" part, narrow and testable).**
1. Metric layer: verified SQL snippets and definitions per database, reused by later questions.
2. Calibration: G2 and V thresholds refit on the harness's own logged traces (labels from execution against gold and
   from J).
3. Verifier fine-tuning: GLiClass fine-tuned on those traces (8-shot per label already gives +0.2 F1 in the paper).
4. Auto-research loop over the harness configuration (section 5).


## 2b. What the tabular and relational models add that the LLM and GLi cannot do

The LLM reads at most a few thousand rows and reasons about them in words; SQL can only count and sum what already
happened; GLi only reads text. None of them can *learn a pattern from millions of rows and apply it*. That is the
tabular / relational FM's job. Seven capabilities, each a tool the LLM calls, each returning a short JSON summary:

| # | capability | example question | why LLM + SQL + GLi cannot | model |
|---|---|---|---|---|
| T1 | predict the future for each entity | "which customers will churn next month?", "expected revenue per store next quarter" | SQL describes the past; the LLM cannot fit a function to 50K+ rows | Kumo Tabular-L (one table); Kumo Relational (many tables, no hand joins) |
| T2 | find the drivers | "what drives churn?", "why did margin drop in region X?" | a few GROUP BYs confuse correlated columns and miss interactions | Kumo Tabular-L + permutation importance / SHAP; TabPFN extensions (CRT p-values) for research |
| T3 | what-if with uncertainty | "if we raise the price 5%, what happens to orders, with a range?" | the LLM guesses; SQL has no counterfactual | partial dependence / ICE on Kumo Tabular-L; quantile outputs give the range |
| T4 | find what does not fit | "which transactions / stores look abnormal?" | nobody can read millions of rows | prediction residuals and outlier scores over all rows (Kumo Tabular-S, fast, many calls) |
| T5 | detect change between periods | "did the customer mix change this quarter? in which columns?" | comparing two periods column by column misses joint shifts | classifier two-sample test: old vs new rows, AUROC = drift, importances = where (Kumo Tabular-S) |
| T6 | similar entities and segments | "find customers like these 20", "what natural groups exist?" | SQL has no notion of similarity across many columns and tables | Kumo Relational graph-layer embeddings (paper 2: best layer, matches hand features with no feature work); the LLM names the segments, GLi labels them |
| T7 | judge the LLM's own hypotheses and features | the LLM says "discounts drive retention" or writes a new SQL feature | the LLM cannot test its own idea on held-out data | held-out signal with and without the feature (the TabFM-Auto mechanism of paper 1; on company-size data the judge is far less noisy than on TabArena's few hundred rows) |

Model roles. **Kumo Tabular-S**: fast probe for many repeated calls (T4, T5, importance loops, interactive answers).
**Kumo Tabular-L**: accuracy for final answers (T1, T2, T3); in paper 1 its plain predictions beat the paper's closed
TabFM on 10 of 17 datasets. **Kumo Relational**: questions that span tables (T1, T6) straight from the database, about
one minute per task; its graph layer is the entity embedding (paper 2). **TabPFN-2.5/3** (research only, non-commercial):
reference numbers and its SHAP / CRT p-value tools. **LightGBM**: a check beside every T1-T3 answer on temporal or very
large tasks, where trees still win (BeyondArena). Limits to respect: Kumo Tabular keeps the first 500 features; Kumo
Relational reads numeric and date columns only (text goes through GLi first, G4), 10 classes, ~20K context rows per call
(the tool samples, time-aware).

How the three work together on one question ("why are we losing customers in the north?"):
GLi links "customers", "north", "losing" to the customer table, region column and churn definition in the metric layer ->
the LLM plans: define churn by SQL, build the task table -> Kumo Relational predicts churn risk per customer (T1) ->
Kumo Tabular-L gives the drivers and what-ifs (T2, T3), LightGBM checks them -> GLi turns the free-text complaint column
into labels that enter as features (G4) -> T5 tests whether the north changed against last year -> the LLM writes
claims, every number recomputed by SQL, every text claim checked by GLi -> the answer lists verified facts with ranges.

Experiments that isolate the tabular part (added to section 3): ablation (E) vs (D) on every benchmark, plus
**C7 capability tasks** where the answer needs T1-T7: RelBench prediction questions (T1), InsightBench planted drivers
and anomalies (T2, T4), DSGym QRData statistical / causal questions (T2, T3), and a planted-drift set built from RelBench
time splits (T5). Without the tabular tool the LLM can only describe; the test is how often the answer is right.

**LLM for the experiments.** The company uses GLM-5.2 inside pi. GLM-5-class models (~744B) need a full 8-GPU node. Plan:
run the harness inside pi (as in paper 1, where pi was the best open harness) with GLM-4.5-Air on one H200 for all
ablations, and repeat the final configuration with GLM-5.2 on an 8-GPU node or through the company endpoint (questions
only, company data stays inside).

## 3. Benchmarks, chosen per claim

| claim | benchmark | why it shows the claim | metric |
|---|---|---|---|
| C1: small-model schema linking + step triage cut SQL errors and tokens | BIRD dev (corrected subset, e.g. the Arcwise / ReViSQL set, plus official dev) ; Spider 2.0-Lite subset (wide schemas) | 81% of errors are schema errors; wide schemas overflow context | execution accuracy, tokens, latency |
| C2: GLiClass is a usable, calibrated verifier | our own trace set from C1 runs ; FC-RewardBench (1,500 correct / incorrect calls) ; LLM-AggreFact (claims) | no published numbers exist: a contribution by itself | AUROC, ECE before / after calibration, latency vs LLM judge |
| C3: the DEEP tool gives deeper, verified insights | InsightBench (100 datasets, planted insights) ; DSGym QRData-Verified and DAEval-Verified ; DABstep | numeric trends over many rows are where LLMs fail | insight score (two judges), exact-match accuracy, tokens |
| C4: relational prediction questions | RelBench 8 entity tasks (our paper-2 setup) ; RelArena | the agent must answer "who will churn" from the database | official test AUROC vs our paper-2 numbers |
| C5: text columns become usable | AutoGluon multimodal text+tabular benchmark (18 datasets) | GLi featurizer vs text embeddings vs LLM labels | AUROC / RMSE, cost |
| C6: long-horizon state | BIRD-Interact Lite ; LongDS-Bench | the ledger vs a growing chat | success rate, early vs late turn accuracy |
| company | 50-100 verified questions on the company database, written with the analysts | practitioners agree public boards do not transfer | accuracy, analyst rating, cost |

## 4. The experiment matrix (one LLM, one budget)

Ablations on each benchmark: (A) LLM alone with SQL tool ; (B) + G1 schema link ; (C) + G2 triage ; (D) + V claim
verification ; (E) + DEEP tool ; (F) full ; (G) full + learning across tasks ; (H) full + auto-research config.
LLMs: Qwen3-Coder-30B-A3B and GLM-4.5-Air (one H200 each, as in paper 1); gpt-oss-120b as a third family; J is the other
family. Report accuracy, tokens, wall time, and the error-type distribution (from G2's own taxonomy, checked by hand on a
sample).

## 5. Auto-research loop (Karpathy-style), with guards

The loop edits the harness configuration (prompts, label sets and thresholds, cascade bands, DEEP-tool settings), runs
the development slice for a fixed budget, keeps a change only if it passes all guards, and logs every attempt.
Guards (from paper 1): development / held-out split fixed before the loop starts, held-out scored once; a change must
beat the paired per-task standard error, not just the mean; an acceptance slice inside development must not get worse;
a fixed experiment budget; corrected benchmark subsets only, because noisy labels reward fitting the noise.

## 6. Order of work

| phase | weeks | content | output |
|---|---|---|---|
| P0 | 1 | verifier study C2: GLiClass vs MiniCheck vs LLM judge vs logprob on BIRD candidate SQL and claims; calibration | first measured numbers for GLiClass as a verifier (short paper on its own) |
| P1 | 1-2 | G4 featurizer C5; DEEP tool as a library function with JSON output | tool + text result |
| P2 | 2-3 | harness (G1, L1, G2, X, L2, V, S) on the open-tabfm-auto runner; BIRD + Spider 2.0-Lite (C1) | ablation table A-F |
| P3 | 2 | InsightBench + DSGym + DABstep (C3); RelBench (C4) | ablation table, insight examples |
| P4 | 1-2 | learning across tasks (G) and auto-research (H) with guards | held-out check of what the loop found |
| P5 | 1 | paper; company pilot on the verified question set (data stays inside the company) | paper 3 |

## 7. Risks
GLiClass may not separate good from bad steps even after calibration (then the triage role is reported as a negative
result and MiniCheck / a fine-tuned GLiClass takes over); judge noise on InsightBench (two judges, agreement reported);
benchmark label errors (corrected subsets); GPU budget (DEEP tool and two LLMs: three H200s at peak, or time-sharing);
TabPFN-2.5+ licences are non-commercial (research only; company path uses TabICLv2 / Kumo Tabular).
