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
