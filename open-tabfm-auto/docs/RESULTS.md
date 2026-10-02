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
