# Auto-generated results tables
_from 15 runs under runs, /home/user/frozen-embeddings-retrieval/runs_

## T1. Frozen backbones with the identity pipeline (3-fold CV, lower is better)
_no backbone-comparison runs yet_

## T2. Pipeline search: harness x LLM x backbone (held-out test error)
| dataset        | backbone   | harness     | llm    | metric   |   P0 test |   P* test |   gain % |   evals |   HistGB |
|:---------------|:-----------|:------------|:-------|:---------|----------:|----------:|---------:|--------:|---------:|
| breast_cancer  | tabpfn     | heuristic   | -      | 1-auroc  |    0.0061 |    0.0061 |   0.0000 |      20 |   0.0108 |
| synth_entities | tabpfn     | heuristic   | -      | 1-auroc  |    0.4676 |    0.4266 |   8.7750 |      20 |   0.4585 |
| synth_physics  | tabpfn     | claude-code | sonnet | rmse     |    0.1108 |    0.0830 |  25.0620 |       4 |   0.3277 |
| synth_physics  | tabpfn     | heuristic   | -      | rmse     |    0.1108 |    0.0852 |  23.1440 |      20 |   0.3277 |
| wine           | tabpfn     | none        | -      | logloss  |    0.0085 |    0.0085 |   0.0000 |       1 |   0.0069 |

## T3. TabArena protocol vs the paper (per-dataset test error, P0 / P* vs TabFM / TabFM-Auto Opus 5)
_no TabArena-protocol runs yet_

## T4. Backbone transfer of discovered pipelines
|                             |     P* |     P0 |   gain % |
|:----------------------------|-------:|-------:|---------:|
| ('synth_physics', 'hgb')    | 0.1813 | 0.3277 |  44.6840 |
| ('synth_physics', 'tabpfn') | 0.0830 | 0.1108 |  25.0620 |

## T5. Retrieval benchmark (synthetic shop DB; mean/std over seeds)
|                                     |   seg P@10 mean |   seg P@10 std |   seg P@10 count |   seg kNN acc mean |   seg kNN acc std |   seg kNN acc count |   fut MAP@10 mean |   fut MAP@10 std |   fut MAP@10 count |
|:------------------------------------|----------------:|---------------:|-----------------:|-------------------:|------------------:|--------------------:|------------------:|-----------------:|-------------------:|
| ('synth_shop', 'agg')               |           0.726 |          0.000 |            5.000 |              0.843 |             0.000 |               5.000 |             0.048 |            0.000 |              5.000 |
| ('synth_shop', 'gnn')               |           0.880 |          0.006 |            5.000 |              0.904 |             0.005 |               5.000 |             0.048 |            0.001 |              5.000 |
| ('synth_shop', 'openrfm')           |           0.249 |          0.000 |            5.000 |              0.306 |             0.000 |               5.000 |             0.032 |            0.000 |              5.000 |
| ('synth_shop', 'openrfm_ctx16')     |           0.237 |        nan     |            1.000 |              0.291 |           nan     |               1.000 |             0.034 |          nan     |              1.000 |
| ('synth_shop', 'openrfm_random')    |           0.168 |          0.000 |            4.000 |              0.169 |             0.000 |               4.000 |             0.013 |            0.000 |              4.000 |
| ('synth_shop', 'row')               |           0.161 |          0.000 |            5.000 |              0.146 |             0.000 |               5.000 |             0.022 |            0.000 |              5.000 |
| ('synth_shop', 'tabpfn_agg')        |           0.588 |          0.340 |            4.000 |              0.651 |             0.312 |               4.000 |             0.039 |            0.013 |              4.000 |
| ('synth_shop', 'tabpfn_agg_churn')  |           0.182 |        nan     |            1.000 |              0.173 |           nan     |               1.000 |             0.022 |          nan     |              1.000 |
| ('synth_shop', 'tabpfn_agg_feat4')  |           0.680 |          0.030 |            3.000 |              0.787 |             0.028 |               3.000 |             0.042 |            0.001 |              3.000 |
| ('synth_shop', 'tabpfn_agg_kmeans') |           0.824 |          0.021 |            3.000 |              0.873 |             0.013 |               3.000 |             0.048 |            0.001 |              3.000 |
| ('synth_shop', 'tabpfn_agg_rand4')  |           0.525 |          0.325 |            3.000 |              0.617 |             0.314 |               3.000 |             0.037 |            0.012 |              3.000 |
| ('synth_shop', 'tabpfn_row')        |           0.171 |          0.005 |            4.000 |              0.174 |             0.012 |               4.000 |             0.021 |            0.001 |              4.000 |

## T6. RelBench rel-hm user-item-purchase (official evaluator)
_no RelBench runs yet_
