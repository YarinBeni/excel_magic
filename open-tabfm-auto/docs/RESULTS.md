# Results

All numbers come from `runs/*/metrics.json`. CPU only (4 cores), frozen TabPFN v2, coding agent = Claude Code with Sonnet.

## Pipeline search (examples/02_pipeline_search.py)

| dataset | metric | P0 (identity) test | P* (agent) test | gain | evals | agent cost |
|---|---|---|---|---|---|---|
| synth_physics (regression, 500 rows, physically named columns) | RMSE | 0.1108 | 0.0830 | -25.1% | 4 | $0.15, 56 s |

Baselines on the raw columns: HistGradientBoosting 0.3277, RandomForest 0.3183. The agent rediscovered the planted
Strouhal-type ratio `frequency * chord / velocity`, its log and thickness-scaled variants (`runs/.../best_pipeline.py`).

## SQL database agent (examples/03_sql_database_agent.py)

Synthetic shop database (customers / products / orders / order_items / support_tickets), task: churn in the 90 days after
the cutoff.

| table | metric | P0 test | P* test | HistGB |
|---|---|---|---|---|
| agent-built table (38 SQL features) | 1-AUROC | 0.2930 | 0.2752 | 0.3145 |
| hand-written reference table (9 features) | 1-AUROC | 0.3165 | - | 0.3615 |

Phase 1 (schema exploration + SQL) cost $0.19, phase 2 (3 evaluations) $0.10.

## Backbone transfer (examples/04_backbone_transfer.py)

| dataset | model | P0 test | P* test | gain |
|---|---|---|---|---|
| synth_physics | tabpfn (v2) | 0.1108 | 0.0830 | +25.1% |
| synth_physics | hgb | 0.3277 | 0.1813 | +44.7% |

The pipeline found for TabPFN also helps gradient boosting (the paper's Section B.3 effect).

## No-LLM heuristic search vs the LLM agent (examples/02_pipeline_search.py --harness heuristic, 20 evals, TabPFN v2)

| dataset | metric | P0 test | heuristic P* test | gain | Sonnet agent P* test | gain | HistGB |
|---|---|---|---|---|---|---|---|
| synth_physics (planted physical ratio) | RMSE | 0.1108 | 0.0852 | -23.1% | 0.0830 | -25.1% | 0.3277 |
| synth_entities (planted entity co-occurrence) | 1-AUROC | 0.4676 | 0.4266 | -8.8% | - | - | 0.4585 |
| breast_cancer (clean, saturated) | 1-AUROC | 0.0061 | 0.0061 | 0.0% | - | - | 0.0108 |

Runs: `runs/20261002T074652Z_heur_synth_physics`, `runs/20261002T075458Z_heur_synth_entities`, `runs/20261002T080519Z_heur_breast_cancer`.
The heuristic (generic crosses of the top-5 informative features, count encoding, winsorizing, more ensemble members)
recovers most of the physics gain without knowing physics: pairwise ratios of the right columns happen to be in its
library. Where the signal needs a formula or domain reading, the LLM's margin should grow; on an already-saturated dataset
neither helps. Cost: 7-10 CPU minutes per dataset, no API.

## TabArena protocol, wave 1 (cluster job J2, 2026-10-02): the 17 datasets with <= 2,500 rows

Frozen **Kumo Tabular-S** (n_estimators=8, H200), **no LLM** (greedy heuristic, 24 evals / <= 90 min), search on fold 0's
training split, P0 and P* scored on all 30 official splits; paper rows from Table 9 (TabFM 400M + Claude Code/Opus 5).

| dataset | metric | our P0 | our P* | our gain | paper TabFM | paper TabFM-Auto | paper gain |
|---|---|---|---|---|---|---|---|
| anneal | logloss | 0.0118 | 0.0108 | +8.5% | 0.0125 | 0.0103 | +17.6% |
| airfoil_self_noise | rmse | 0.9793 | 0.9161 | +6.5% | 1.0734 | 0.9165 | +14.6% |
| hazelnut-spread-contaminant-detection | 1-auroc | 0.0048 | 0.0047 | +1.3% | 0.0023 | 0.0021 | +8.7% |
| Fitness_Club | 1-auroc | 0.1787 | 0.1774 | +0.7% | 0.1789 | 0.1787 | +0.1% |
| website_phishing | logloss | 0.2144 | 0.2134 | +0.4% | 0.2104 | 0.2067 | +1.8% |
| QSAR_fish_toxicity | rmse | 0.8547 | 0.8529 | +0.2% | 0.8535 | 0.8483 | +0.6% |
| blood-transfusion-service-center | 1-auroc | 0.2458 | 0.2453 | +0.2% | 0.2441 | 0.2431 | +0.4% |
| Marketing_Campaign | 1-auroc | 0.0660 | 0.0659 | +0.1% | 0.0732 | 0.0616 | +15.8% |
| credit-g | 1-auroc | 0.1979 | 0.1983 | -0.2% | 0.1944 | 0.1940 | +0.2% |
| Another-Dataset-on-used-Fiat-500 | rmse | 713.9 | 719.0 | -0.7% | 703.3 | 693.5 | +1.4% |
| qsar-biodeg | 1-auroc | 0.0596 | 0.0603 | -1.2% | 0.0580 | 0.0582 | -0.3% |
| healthcare_insurance_expenses | rmse | 4553.6 | 4613.9 | -1.3% | 4417.6 | 4344.1 | +1.7% |
| concrete_compressive_strength | rmse | 3.8268 | 3.8782 | -1.3% | 3.9666 | 3.8169 | +3.8% |
| diabetes | 1-auroc | 0.1596 | 0.1623 | -1.7% | 0.1580 | 0.1471 | +6.9% |
| maternal_health_risk | logloss | 0.3847 | 0.3952 | -2.7% | 0.3706 | 0.3648 | +1.6% |
| MIC | logloss | 0.4318 | 0.4254 | +1.5% | 0.4282 | 0.4178 | +2.4% |
| Is-this-a-good-customer | 1-auroc | 0.2474 | 0.2532 | -2.3% | 0.2466 | 0.2432 | +1.4% |

Summary (17/17): mean relative gain +0.5% (paper's Opus 5 agent: +4.6% on the same 17); P* beats P0 on 9/17; the frozen
open model's identity pipeline is within a few percent of the paper's TabFM on 16/17 (better on 5; the exception is
hazelnut, where TabFM is 2x better). The no-LLM search captures a third to a half of the paper's gain where the signal
is generic (anneal, airfoil) and overfits fold-0 CV on small noisy tables (diabetes, maternal_health_risk), the same
regime where the paper reports its own small losses (credit-g). Runs: `reports/runs/*J2_heuristic_*` on the cluster branch.

## Frozen open backbones, identity pipeline (cluster job J1b, H200, 3-fold CV, lower is better)

| dataset | metric | TabPFN v2 | TabPFN-2.5 | TabICLv2 | Kumo-S | Kumo-M | Kumo-L | HistGB |
|---|---|---|---|---|---|---|---|---|
| wine | logloss | 0.0356 | 0.0361 | **0.0158** | 0.0294 | 0.0186 | 0.0175 | 0.0965 |
| breast_cancer | 1-auroc | 0.0042 | 0.0039 | 0.0042 | 0.0039 | 0.0038 | **0.0033** | 0.0059 |
| diabetes | rmse | 54.28 | 54.18 | **53.73** | 54.59 | 54.20 | 54.52 | 60.52 |
| synth_physics | rmse | 0.0920 | 0.0961 | 0.0951 | 0.0868 | 0.0854 | **0.0838** | 0.3539 |
| synth_entities | 1-auroc | 0.5010 | 0.4492 | 0.4439 | 0.4424 | **0.4213** | 0.4319 | 0.4873 |

Every 3-fold CV took 0.4–3.7 s on the H200 (Kumo-L the slowest). Kumo Tabular-L is best on 3/5, TabICLv2 on 2/5; TabPFN v2
(the only CPU-mirror model) is last among the TFMs on 4/5. EXAONE-Tabular failed on a NumPy-input requirement (fixed).

## Backbone transfer of the pipelines found for TabPFN v2 (J1b, held-out test)

| dataset | pipeline found by | TabPFN v2 | TabPFN-2.5 | TabICLv2 | Kumo-S | Kumo-M | Kumo-L |
|---|---|---|---|---|---|---|---|
| synth_physics | Sonnet (Strouhal ratio) | +15.5% | -11.3% | +4.2% | -1.1% | +0.8% | -1.1% |
| synth_entities | heuristic (count features) | +11.4% | +13.6% | -0.5% | +26.0% | -18.2% | +5.3% |

Unlike the paper's finding (pipelines found for TabFM help every weaker model), pipelines found for the *weakest* model
transfer unevenly to stronger ones: the ratio feature that TabPFN needs is largely redundant for Kumo, and the count
features help Kumo-S a lot but hurt Kumo-M. The search should be run per backbone, or around the strongest one.

## RelBench rel-hm user-item-purchase (cluster job J7, 2026-10-02, official evaluator, MAP@12 x100)

| method | val | test |
|---|---|---|
| GlobalPopularity | 0.342 | 0.292 |
| PastVisit | 1.904 | 2.199 |
| kNN-CF[row] | 0.136 | 0.144 |
| kNN-CF[agg] | 0.258 | 0.259 |
| kNN-CF[tabpfn_agg_kmeans] | 0.248 | 0.227 |
| kNN-CF[tabpfn_agg_random] | 0.207 | 0.170 |
| kNN-CF[tabpfn_row_kmeans] | 0.135 | 0.140 |
| PastVisit+kNN-CF[any embedding] | 1.86-1.89 | 2.16-2.19 |
| kNN-CF[purchase_matrix] (user-kNN, no model) | 1.162 | 1.200 |
| PastVisit+kNN-CF[purchase_matrix] | 1.919 | 2.216 |
| ItemKNN | 1.035 | 1.073 |
| PastVisit+ItemKNN | 1.901 | 2.197 |

Published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM 2.73, ContextGNN 2.93.
Negative result for the frozen-embedding hypothesis at item level, and attributed: the same kNN scoring over the raw
purchase matrix reaches 1.20 and is the only hybrid above PastVisit, so the customer-level embeddings (not the retrieval
step) lack the item signal. J7d feeds SVD factors of the purchase matrix through the frozen TabPFN and compares with the
raw factors.
Source: `reports/runs/20261002T160111Z_J7_hm_full/` in the cluster branch.

## Open LLM via vLLM: Qwen3-Coder-30B-A3B-Instruct (cluster jobs J3 50375 / J5 50402, 2026-10-02)

Frozen Kumo Tabular-S (n_estimators=4; 8 for TabArena), OpenAI-compatible tool loop (`--harness openai`) over vLLM 0.16,
budget 16 evals / 40 min (TabArena-Lite: 24 / 60), held-out test, two independent runs.

| dataset | metric | P0 J3 | P* J3 | gain J3 | P0 J5 | P* J5 | gain J5 |
|---|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0869 | 0.0857 | +1.3% | 0.0865 | 0.0840 | +2.9% |
| synth_entities | 1-auroc | 0.4387 | 0.4437 | -1.1% | 0.4545 | 0.4499 | +1.0% |
| breast_cancer | 1-auroc | 0.0051 | 0.0053 | -2.9% | 0.0044 | 0.0069 | -56.7% |
| airfoil_self_noise (Lite r0f0) | rmse | 0.8209 | 0.8282 | -0.9% | 0.8277 | 0.8088 | +2.3% |
| blood-transfusion (Lite) | 1-auroc | 0.2647 | 0.2582 | +2.5% | 0.2631 | 0.2655 | -0.9% |
| credit-g (Lite) | 1-auroc | 0.2056 | 0.2062 | -0.3% | 0.2054 | 0.2062 | -0.4% |

Agent stats per search: 16-41 turns, ~100-310k input tokens, 0.6-7k output tokens, 90-134 s wall on one H200. One run
called `run_eval` 38 times on an unchanged pipeline (harness now refuses that). Runs: `reports/runs/*J3_openai_*`,
`reports/runs/*J5_Qwen3-Coder-30B-A3B-Instruct_*` on the cluster branch.

## CLI harness: aider + Qwen3-Coder-30B-A3B via vLLM (cluster job J6 task 2, 50377, 2026-10-02)

Frozen Kumo Tabular-S (n_estimators=4), `--harness cli --agent-cmd <aider loop>` (8 rounds of aider --message-file
-> tabfm-eval -> score appended to the brief), budget 16 evals / 40 min, held-out test.

| dataset | metric | P0 | P* | gain | evals | HistGB |
|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0861 | 0.0880 | -2.2% | 9 | 0.324 |
| synth_entities | 1-auroc | 0.3842 | 0.2239 | **+41.7%** | 9 | 0.473 |
| breast_cancer | 1-auroc | 0.0051 | 0.0072 | -40.0% | 9 | 0.011 |

The synth_entities pipeline (`reports/runs/*J6_aider_synth_entities/best_pipeline.py`): role x resource pair counts,
manager / role / resource counts over train+test, their ratios and logs, a label-encoded triple, 4 context views
(all, class-balanced, random, minority-oversampled), sigmoid sharpening in postprocess. CV 0.470 -> 0.317 at eval 2,
best 0.296 at eval 5. The same LLM in our tool loop (J3/J5) never found the frequency encodings.

## CLI harness: Qwen Code 0.24.7 + Qwen3-Coder-30B-A3B via vLLM (cluster job J6 task 1, 50428, 2026-10-02)

Same setup as the aider row (frozen Kumo Tabular-S, 16 evals / 40 min). First attempt (50408) failed on every turn because
Qwen Code requests max_tokens = the full context and the server was 32k; served at 128k it works.

| dataset | metric | P0 | P* | gain | evals | agent |
|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0858 | 0.0849 | +1.0% | 9 | 138 s |
| synth_entities | 1-auroc | 0.4592 | 0.3422 | **+25.5%** | 9 | 143 s, 22 API calls, 616k prompt tokens |
| breast_cancer | 1-auroc | 0.0045 | 0.0053 | -16.1% | 6 | 96 s |

Entity task, all with Qwen3-Coder-30B and the same frozen backbone: aider loop +41.7% > Qwen Code +25.5% ~ no-LLM
heuristic +25-29% > our OpenAI tool loop -1..+1%.

## Open LLM via vLLM: gpt-oss-20b (cluster job J5 task 3, 50431, 2026-10-02)

Same setup as the Qwen3-Coder rows (frozen Kumo Tabular-S, our OpenAI tool loop, 16 evals / 40 min).

| dataset | metric | P0 | P* | gain | evals |
|---|---|---|---|---|---|
| synth_physics | rmse | 0.0864 | 0.0875 | -1.3% | 2 |
| synth_entities | 1-auroc | 0.4307 | 0.4726 | -9.7% | 6 |
| breast_cancer | 1-auroc | 0.0055 | 0.0051 | +7.9% | 6 |

vLLM 0.16's harmony tool parser leaked channel markers into the tool names (`run_eval<|channel|>commentary`), which the
harness counted as unknown tools; names are now normalised. The TabArena-Lite run failed on an empty response
(`choices=None`), now guarded; task resubmitted.

## RelBench rel-hm, interaction factors through the frozen TFM (cluster job J7d, 50399, official evaluator, MAP@12 x100)

| method | val | test |
|---|---|---|
| kNN-CF[purchase_matrix] (sparse cosine, no model) | 1.162 | 1.200 |
| kNN-CF[svd] (64 SVD factors of log1p counts) | 0.822 | 0.833 |
| kNN-CF[svd_agg] (factors + customer aggregates) | 0.661 | 0.654 |
| kNN-CF[tabpfn_svd_kmeans] (frozen TabPFN over svd_agg, k-means target) | 0.401 | 0.419 |
| kNN-CF[tabpfn_svd_random] | 0.514 | 0.475 |
| kNN-CF[agg] / [tabpfn_agg_kmeans] | 0.258 / 0.248 | 0.259 / 0.227 |
| best hybrid: Past+kNN-CF[purchase_matrix] | 1.919 | 2.216 |

The frozen TFM hidden state keeps about half the retrieval value of its input; the k-means pseudo-target does not help
here. Negative at item level, closed. Source: `reports/runs/20261002T173110Z_J7d_hm_svd_full/`.

## gpt-oss-20b rerun with the fixed harness (J5 task 3, 50434)

| dataset | metric | P0 | P* | gain | evals |
|---|---|---|---|---|---|
| synth_physics | rmse | 0.0869 | 0.0834 | +4.1% | 6 |
| synth_entities | 1-auroc | 0.3803 | 0.4144 | -8.9% | 9 (33 writes) |
| breast_cancer | 1-auroc | 0.0055 | 0.0060 | -7.9% | 10 |
| airfoil_self_noise (Lite) | rmse | 0.8173 | 0.8007 | +2.0% | |
| blood-transfusion (Lite) | 1-auroc | 0.2627 | 0.2657 | -1.1% | |
| credit-g (Lite) | 1-auroc | 0.2071 | 0.2076 | -0.2% | |

## Open LLM via vLLM: GLM-4.5-Air-FP8 (cluster job J5 task 2, 50414, 2026-10-02)

Same setup (frozen Kumo Tabular-S, our OpenAI tool loop, 16 evals / 40 min; TabArena-Lite 24 / 60). Served FP8 at 0.90
GPU memory on one H200 (bf16 does not fit).

| dataset | metric | P0 | P* | gain | evals | tokens in / out | wall |
|---|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0884 | 0.0839 | +5.1% | 16 | 471k / 22k | 293 s |
| synth_entities | 1-auroc | 0.3972 | 0.2240 | **+43.6%** | 16 | 523k / 28k | 350 s |
| breast_cancer | 1-auroc | 0.0053 | 0.0073 | -38.9% | 9 | 301k / 26k | 306 s |
| airfoil_self_noise (Lite) | rmse | 0.8264 | 0.8337 | -0.9% | | | |
| blood-transfusion (Lite) | 1-auroc | 0.2639 | 0.2602 | +1.4% | | | |
| credit-g (Lite) | 1-auroc | 0.2069 | 0.2039 | +1.4% | | | |

Best open LLM in our tool loop so far; the entity pipeline uses group-by frequency features like aider's.

## CLI harness: pi + Qwen3-Coder-30B-A3B via vLLM (cluster job J6 task 0, 50515, 2026-10-02)

Same setup as the aider / Qwen Code rows. First attempt (50377_0) died creating the Node env in a race with task 1.

| dataset | metric | P0 | P* | gain | evals | agent |
|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0880 | 0.0888 | -0.9% | 10 | 108 s |
| synth_entities | 1-auroc | 0.4515 | 0.2242 | **+50.3%** | 6 | 89 s |
| breast_cancer | 1-auroc | 0.0051 | 0.0060 | -17.1% | 12 | 146 s |

Entity task, one frozen backbone, one open LLM: pi +50.3% > aider +41.7% > Qwen Code +25.5% ~ heuristic +25-29% > our
tool loop ~0. All CLI agents found the same role x resource frequency encodings.

## Open LLM via vLLM: Qwen3-32B (cluster job J5 task 1, 50405, 2026-10-02)

Same setup (frozen Kumo Tabular-S, our OpenAI tool loop, `--reasoning-parser qwen3`, 40k context).

| dataset | metric | P0 | P* | gain | evals | tokens in / out | wall |
|---|---|---|---|---|---|---|---|
| synth_physics | rmse | 0.0862 | 0.0863 | -0.2% | 1 | 3.5k / 4.4k | 84 s (stopped after describe_data) |
| synth_entities | 1-auroc | 0.4546 | 0.4718 | -3.8% | 16 | 830k / 88k | 1789 s |
| breast_cancer | 1-auroc | 0.0051 | 0.0053 | -2.9% | 6 | 114k / 52k | 1034 s |
| airfoil_self_noise (Lite) | rmse | 0.8318 | 0.8261 | +0.7% | | | |
| blood-transfusion (Lite) | 1-auroc | 0.2656 | 0.2639 | +0.6% | | | |
| credit-g (Lite) | 1-auroc | 0.2062 | 0.2055 | +0.3% | | | |

Weakest open model in the sweep: long reasoning, many rewrites, no useful features.

## TabArena wave 1 with LLM agents (cluster job J8, 2026-10-02): pi + Qwen3-Coder (17/17), GLM-4.5-Air tool loop (6/17)

Same protocol as the heuristic wave (frozen Kumo Tabular-S n_estimators=8, search on r0f0 train, P0/P* on all 30 official
splits), 16 evals / 40 min per dataset. pi: mean gain -0.94%, median -0.02%, P* < P0 on 9/17; P* below the paper's
TabFM-Auto on 15/17 and below the paper's TabFM on 13/17. Heuristic (J2): +0.47% mean. Paper (Opus 5): +4.62%.

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

Runs: `reports/runs/*J8_pi_qwen3coder_*`, `reports/runs/*J8_openai_glm45air_*`.

## RelBench rel-hm user-churn, entity-level kNN probe (cluster job J12, 50618, 2026-10-02)

20,000 customers at the last train timestamp (churn rate 0.817), 365-day history, 5 random folds over customers, kNN
probe k = 50 (mean neighbour label), AUROC. Not the official temporal test split.

| method | AUROC |
|---|---|
| Supervised HGB on aggregates | 0.673 |
| kNN[agg] | 0.653 |
| kNN[tabpfn_agg_kmeans] | 0.648 |
| kNN[tabpfn_svd_kmeans] | 0.645 |
| kNN[tabpfn_agg_random] | 0.643 |
| kNN[svd_agg] | 0.638 |
| kNN[svd] | 0.590 |
| kNN[row] | 0.520 |
| majority prior | 0.500 |

Supervised TabPFN on the aggregates (labels in context, J12c): 0.672, equal to HGB. Frozen embedding = its input, 2.4 points below the same model used with labels.
Source: `reports/runs/20261002T203451Z_J12_churn_full/`.

## Rich-context tool loop (cluster job J13, 50619): Qwen3-Coder-30B with the CLI agents' context in the first message

`TABFM_LOOP_RICH=1`: file listing + pipeline.py + data summary prepended to the brief; otherwise identical to J3/J5.

| dataset | metric | P0 | P* | gain | evals |
|---|---|---|---|---|---|
| synth_physics | rmse | 0.0860 | 0.0866 | -0.7% | 8 |
| synth_entities | 1-auroc | 0.4440 | 0.4658 | -4.9% | 10 |
| breast_cancer | 1-auroc | 0.0050 | 0.0061 | -23.5% | 12 |
| airfoil_self_noise (Lite) | rmse | 0.8222 | 0.7953 | +3.3% | |
| blood-transfusion (Lite) | 1-auroc | 0.2650 | 0.2639 | +0.4% | |
| credit-g (Lite) | 1-auroc | 0.2071 | 0.2043 | +1.4% | |

No improvement over the plain loop on the synthetic tasks (pi with the same LLM: +50% on entities): the harness gap is
not the starting information.

## Repeated-CV judge on small tables (cluster job J11, 2026-10-02): 3 x 3-fold vs 1 x 3-fold, same budgets

Frozen Kumo Tabular-S; TabArena protocol on the five wave-1 datasets with < 1000 training rows, and the three bundled
tasks. Gains are relative error reduction P* vs P0 on the official splits (held-out test for the bundled tasks).

| dataset | heuristic 1x3 (J2) | heuristic 3x3 (J11) | pi 1x3 (J8) | pi 3x3 (J11) |
|---|---|---|---|---|
| anneal | +8.5% | +10.5% | -5.2% | (running) |
| diabetes | -1.7% | +1.2% | +0.8% | +0.6% |
| QSAR_fish_toxicity | +0.2% | +0.2% | -0.1% | -0.0% |
| blood-transfusion | +0.2% | -0.6% | -0.1% | -0.1% |
| credit-g | -0.2% | -0.6% | -2.2% | -2.4% |
| synth_physics | +0.6% (J1) | -1.0% | -0.9% (J6) | -0.1% |
| synth_entities | +25-29% (J1) | +8.3% | +50.3% (J6) | +28.1% |
| breast_cancer | 0% | 0% | -17.1% (J6) | -5.4% |

The repeated judge removes most of the small-table loss (pi breast_cancer -17% -> -5%) and lifts the heuristic on
anneal/diabetes; the synthetic entity gains are smaller under the stricter judge (fewer candidates pass).

## Budget 64 on the paper's high-gain datasets (cluster job J10, 2026-10-02) and split transfer

Frozen Kumo Tabular-S, 64 evals / 150 min, `--cv-repeats auto`; gains = relative error reduction over all 30 official
splits. Standard-budget rows from J2 (heuristic, 24 evals), J8 (pi / GLM, 16 evals).

| dataset | heuristic 24 | heuristic 64 | pi 16 | pi 64 | GLM 16 | GLM 64 | paper Opus 5 |
|---|---|---|---|---|---|---|---|
| anneal | +8.5% | +5.7% | -5.2% | -4.6% | -2.7% | -0.8% | +17.6% |
| Marketing_Campaign | +0.1% | -0.7% | -14.2% | -15.3% | | -47.7% | +15.8% |
| airfoil_self_noise | +6.5% | +4.8% | +3.7% | +2.5% | +1.8% | +4.6% | +14.6% |
| hazelnut | +1.3% | -4.7% | +0.2% | -4.0% | | -1.6% | +8.7% |
| diabetes | -1.7% | +0.5% | +0.8% | -0.0% | +1.3% | +3.3% | +6.9% |
| mean | +2.9% | +1.1% | -2.9% | -4.3% | +0.1% | -8.4% | +12.7% |

Evaluations actually used at budget 64: heuristic 31-44, pi 8-14, GLM 13-22.

Split transfer (mean over datasets of the gain on the search split's own test r0f0 vs the mean over the other 29 splits,
and the fraction of splits where P* beats P0):

| setup | gain on r0f0 | gain on other 29 | splits improved |
|---|---|---|---|
| heuristic 24 | -0.5% | -0.1% | 0.51 |
| heuristic 64 | +0.8% | +0.3% | 0.57 |
| pi 16 | +1.4% | -1.3% | 0.50 |
| pi 64 | -0.1% | -6.4% | 0.39 |
| GLM 16 | -1.3% | -1.2% | 0.50 |
| GLM 64 | -7.3% | -11.5% | 0.54 |

A candidate selected by 3-fold CV on one split of a 500-2500-row table is as likely to hurt as to help on another split
of the same data: the search is selecting CV noise. Runs: `reports/runs/*J10_*`, per-split rows in each `results.csv`.

## Per-backbone TabArena wave 1 (cluster job J9, 2026-10-02): heuristic search around Kumo Tabular-L and TabICLv2

Same protocol as J2 (24 evals / 90 min, search on r0f0 train, P0 / P* on all 30 official splits).

| dataset | metric | Kumo-S P0 / P* | Kumo-L P0 / P* | TabICLv2 P0 / P* | paper TabFM / TabFM-Auto |
|---|---|---|---|---|---|
| Another-Dataset-on-used-Fiat-500 | rmse | 713.9 / 719 | 703.9 / 696.2 | 715.2 / 712.7 | 703.3 / 693.5 |
| Fitness_Club | 1-auroc | 0.1787 / 0.1774 | 0.1784 / 0.178 | 0.1787 / 0.178 | 0.1789 / 0.1787 |
| Is-this-a-good-customer | 1-auroc | 0.2474 / 0.2532 | 0.2446 / 0.2534 | 0.2522 / 0.2557 | 0.2466 / 0.2432 |
| MIC | logloss | 0.4318 / 0.4254 | 0.4187 / 0.4177 | 0.4445 / 0.4392 | 0.4282 / 0.4178 |
| Marketing_Campaign | 1-auroc | 0.06597 / 0.06588 | 0.06401 / 0.06356 | 0.06659 / 0.06657 | 0.0732 / 0.0616 |
| QSAR_fish_toxicity | rmse | 0.8547 / 0.8529 | 0.8565 / 0.8529 | 0.8584 / 0.8585 | 0.8535 / 0.8483 |
| airfoil_self_noise | rmse | 0.9793 / 0.9161 | 0.8913 / 0.8594 | 1.076 / 1.007 | 1.073 / 0.9165 |
| anneal | logloss | 0.01183 / 0.01083 | 0.01101 / 0.01198 | 0.01795 / 0.02737 | 0.0125 / 0.0103 |
| blood-transfusion-service-center | 1-auroc | 0.2458 / 0.2453 | 0.2475 / 0.2495 | 0.2446 / 0.2467 | 0.2441 / 0.2431 |
| concrete_compressive_strength | rmse | 3.827 / 3.878 | 3.802 / 3.808 | 3.965 / 4.02 | 3.967 / 3.817 |
| credit-g | 1-auroc | 0.1979 / 0.1983 | 0.1925 / 0.1957 | 0.2031 / 0.206 | 0.1944 / 0.194 |
| diabetes | 1-auroc | 0.1596 / 0.1623 | 0.1564 / 0.1592 | 0.1611 / 0.1616 | 0.158 / 0.1471 |
| hazelnut-spread-contaminant-detection | 1-auroc | 0.004792 / 0.004729 | 0.002673 / 0.002811 | 0.005089 / 0.005089 | 0.0023 / 0.0021 |
| healthcare_insurance_expenses | rmse | 4554 / 4614 | 4548 / 4535 | 4451 / 4451 | 4418 / 4344 |
| maternal_health_risk | logloss | 0.3847 / 0.3952 | 0.3763 / 0.3868 | 0.3986 / 0.3985 | 0.3706 / 0.3648 |
| qsar-biodeg | 1-auroc | 0.0596 / 0.06031 | 0.059 / 0.05937 | 0.05833 / 0.05836 | 0.058 / 0.0582 |
| website_phishing | logloss | 0.2144 / 0.2134 | 0.2096 / 0.2082 | 0.2228 / 0.2216 | 0.2104 / 0.2067 |

| backbone | mean gain P*/P0 | wins | P0 beats paper TabFM | P0 vs paper TabFM (mean rel.) | P* beats paper TabFM-Auto |
|---|---|---|---|---|---|
| Kumo Tabular-S (J2) | +0.5% | 9/17 | 5/17 | -5.8% | 2/17 |
| Kumo Tabular-L | -1.1% | 8/17 | 10/17 | +1.6% | 4/17 |
| TabICLv2 | -2.9% | 7/17 | 3/17 | -10.9% | 1/17 |

GLM-4.5-Air tool loop, final 17/17 (J8): mean -1.4%, 8/17 wins. J11 pi cv3 final: anneal -0.5% (was -5.2% with one repeat).
Runs: `reports/runs/*J9_heuristic_kumoL_*`, `*J9_heuristic_tabiclv2_*`.

## Kumo Relational on rel-hm user-churn (cluster job J14, 50826): entity-level kNN probe, AUROC

| method | AUROC |
|---|---|
| Supervised HGB on aggregates | 0.673 |
| Supervised TabPFN on aggregates | 0.672 |
| kNN[kumo_relational_random] | **0.660** |
| kNN[agg] | 0.653 |
| kNN[tabpfn_agg_kmeans] | 0.648 |
| kNN[kumo_relational_kmeans] | 0.644 |

Graph: customers <- transactions -> articles, 365-day window, 2 hops, 64 context rows; 172 s for 20k customers on one H200.

## Kumo Relational on rel-hm user-item-purchase (J14, 50826, official evaluator, MAP@12 x100)

| method | val | test |
|---|---|---|
| kNN-CF[kumo_relational_kmeans] | 0.201 | 0.184 |
| Past+kNN-CF[kumo_relational_kmeans] | 1.881 | 2.181 |
| kNN-CF[agg] | 0.258 | 0.259 |
| kNN-CF[purchase_matrix] | 1.164 | 1.203 |
| PastVisit | 1.904 | 2.199 |

Random-target variant (J14b, 50921): kNN-CF[kumo_relational_random] 0.253 / 0.244 (val / test), level with the aggregates; Past+ hybrid 2.182. Retrieval study complete.

## [WITHDRAWN] Final-pick selection rules, heuristic searches (J15, 50827): scored with TabPFN v2 by mistake (nested-config bug); redone as J15c

Rules applied to the recorded per-fold CV scores of the J2 searches (Kumo Tabular-S, 24 evals): best = CV-best (paper),
gatedZ = CV-best among candidates whose paired per-fold improvement over P0 exceeds Z standard errors (else P0),
ens3 = average of the 3 best-CV candidates' predictions.

| dataset | best | gated1 | gated2 | ens3 |
|---|---|---|---|---|
| anneal | +3.7 | +3.7 | -1.4 | **+11.2** |
| MIC | +10.2 | +10.2 | +10.2 | +11.3 |
| airfoil_self_noise | +3.5 | +3.5 | +3.5 | +4.1 |
| website_phishing | +2.2 | +2.2 | +2.2 | +2.2 |
| Fitness_Club | +0.2 | +0.2 | +0.2 | +0.2 |
| QSAR_fish_toxicity | +0.1 | +0.1 | +0.1 | +0.1 |
| Another-Dataset-on-used-Fiat-500 | +0.1 | 0.0 | 0.0 | +0.1 |
| credit-g | 0.0 | 0.0 | 0.0 | +0.1 |
| blood-transfusion / hazelnut | 0.0 | 0.0 | 0.0 | 0.0 |
| healthcare_insurance_expenses | -0.9 | -0.9 | 0.0 | -0.3 |
| concrete_compressive_strength | -1.0 | -1.0 | 0.0 | -0.2 |
| maternal_health_risk | -1.6 | +0.5 | +0.5 | -1.7 |
| Is-this-a-good-customer | -1.7 | -1.7 | 0.0 | -1.7 |
| Marketing_Campaign | -1.8 | -1.8 | -1.8 | -1.6 |
| qsar-biodeg | -1.8 | -1.8 | +0.6 | -1.8 |
| diabetes | -3.1 | -3.1 | -3.1 | -3.1 |
| **mean** | **+0.5** | **+0.6** | **+0.7** | **+1.1** |
| datasets losing > 1% | 6 | 5 | 3 | 5 |
| splits where the pick beats P0 | 41% | 42% | 41% | 46% |

Source: `reports/runs/*J2_heuristic_*/selection_rules.csv`.

## Backbone determinism (cluster job J16, 51089)

Identity pipeline, Kumo Tabular-S (n_estimators=8, cuda), three calls per split in one process, then a second process,
then the J8-era library:

| split | process 1 | process 2 | J8-era library | J2 run | J8 run |
|---|---|---|---|---|---|
| anneal r0f0 (logloss) | 0.0176 0.0177 0.0195 | 0.0192 0.0174 0.0174 | 0.0172 0.0178 0.0177 | 0.0184 | 0.0181 |
| anneal r0f2 | 0.0035 0.0047 0.0045 | 0.0031 0.0051 0.0073 | 0.0055 0.0036 0.0045 | 0.0044 | 0.0041 |
| airfoil r0f0 (rmse) | 0.822 0.822 0.833 | 0.825 0.817 0.816 | 0.834 0.828 0.826 | 0.830 | 0.822 |
| airfoil r0f1 | 1.015 1.018 1.027 | 1.007 1.012 1.014 | 1.014 1.004 1.031 | 1.018 | 1.026 |

The backbone is nondeterministic at the +-1% (RMSE) to +-30% (small-fold logloss) level; the runner now seeds torch
per evaluation. J15's P0 (anneal 0.023 / 0.0054 / 0.0031, airfoil 0.975 / 1.156 / 1.046) was TabPFN v2, not Kumo-S.

## Final-pick selection rules, corrected (J15c, 51092): heuristic searches re-scored with Kumo Tabular-S, gain over P0 %

| rule | mean | median | wins | losses > 1% | splits where the pick beats P0 |
|---|---|---|---|---|---|
| best (CV-best, paper) | +0.25 | 0.00 | 7/17 | 6 | 44% |
| gated 1 s.e. | +0.49 | 0.00 | 8/17 | 5 | 47% |
| gated 2 s.e. | +0.29 | 0.00 | 8/17 | 3 | 44% |
| top-3 ensemble | +0.25 | -0.11 | 7/17 | 6 | 44% |

Per dataset (best / gated1 / gated2 / ens3): airfoil +6.1/+6.1/+6.1/+6.2, anneal +7.0/+7.0/+0.2/+7.0, MIC +1.5 all,
Fitness_Club +0.7, website_phishing +0.5, maternal_health_risk -2.8/+0.6/+0.6/-2.7, Is-this-a-good-customer -2.3/-2.3/0.0/-2.3,
diabetes -1.9 all, qsar-biodeg -1.4, concrete -1.3, healthcare -1.3/-1.3/0.0/-1.3, the rest within +-0.7.
The re-scored P0 agrees with J2's P0 within +-2% (backbone nondeterminism); the J2 run's own +0.47% for the CV-best pick
is +0.25% here for the same reason.

## Final-pick selection rules, corrected (J15c, 51106): pi + Qwen3-Coder searches re-scored with Kumo Tabular-S

| rule | mean gain % | wins | losses > 1% | splits where the pick beats P0 |
|---|---|---|---|---|
| best | -0.95 | 6/17 | 3 | 40% |
| gated 1 s.e. | -0.68 | 4/17 | 2 | 26% |
| gated 2 s.e. | -0.63 | 1/17 | 1 | 8% |
| top-3 ensemble | -0.80 | 7/17 | 3 | 49% |

Marketing_Campaign -14.6% under every rule (the candidate is significantly better in CV and worse on all other splits);
airfoil +3.8%, anneal -4.8% (gated: 0), credit-g -2.2%. Re-scored P0 within 3.7% of J8's.
