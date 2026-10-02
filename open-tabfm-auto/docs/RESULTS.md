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
| PastVisit+kNN-CF[any] | 1.897 | 2.191 |

Published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM 2.73, ContextGNN 2.93.
Negative result for the frozen-embedding hypothesis at item level: every kNN row is below the popularity prior and the
hybrids never improve on PastVisit. Reference rows (user-kNN on the raw purchase matrix, item-kNN) are queued (J7c).
Source: `reports/runs/20261002T160111Z_J7_hm_full/` in the cluster branch.
