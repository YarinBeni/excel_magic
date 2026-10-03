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
