# Auto-generated results tables
_from 45 runs under runs, /home/user/excel_magic/reports/runs, /home/user/frozen-embeddings-retrieval/runs_

## T1. Frozen backbones with the identity pipeline (3-fold CV, lower is better)
|                               |     hgb |   kumo-tabular-l |   kumo-tabular-m |   kumo-tabular-s |   tabicl |   tabpfn |   tabpfn-2.5 |
|:------------------------------|--------:|-----------------:|-----------------:|-----------------:|---------:|---------:|-------------:|
| ('breast_cancer', '1-auroc')  |  0.0059 |           0.0033 |           0.0038 |           0.0039 |   0.0042 |   0.0042 |       0.0039 |
| ('diabetes', 'rmse')          | 60.5173 |          54.5179 |          54.1992 |          54.5932 |  53.7263 |  54.2821 |      54.1759 |
| ('synth_entities', '1-auroc') |  0.4873 |           0.4319 |           0.4213 |           0.4424 |   0.4439 |   0.5010 |       0.4492 |
| ('synth_physics', 'rmse')     |  0.3539 |           0.0838 |           0.0854 |           0.0868 |   0.0951 |   0.0920 |       0.0961 |
| ('wine', 'logloss')           |  0.0965 |           0.0175 |           0.0186 |           0.0294 |   0.0158 |   0.0356 |       0.0361 |

## T2. Pipeline search: harness x LLM x backbone (held-out test error)
| dataset        | backbone       | harness     | llm    | metric   |   P0 test |   P* test |   gain % |   evals |   HistGB |
|:---------------|:---------------|:------------|:-------|:---------|----------:|----------:|---------:|--------:|---------:|
| breast_cancer  | tabpfn         | heuristic   | -      | 1-auroc  |    0.0061 |    0.0061 |   0.0000 |      20 |   0.0108 |
| synth_entities | kumo-tabular-s | heuristic   | -      | 1-auroc  |    0.4623 |    0.3281 |  29.0230 |      20 |   0.4733 |
| synth_entities | kumo-tabular-s | heuristic   | -      | 1-auroc  |    0.4412 |    0.3327 |  24.5890 |      20 |   0.4733 |
| synth_entities | tabpfn         | heuristic   | -      | 1-auroc  |    0.4676 |    0.4266 |   8.7750 |      20 |   0.4585 |
| synth_physics  | kumo-tabular-s | heuristic   | -      | rmse     |    0.0876 |    0.0872 |   0.5530 |      20 |   0.3238 |
| synth_physics  | kumo-tabular-s | heuristic   | -      | rmse     |    0.0872 |    0.0872 |   0.0350 |      20 |   0.3238 |
| synth_physics  | tabpfn         | claude-code | sonnet | rmse     |    0.1108 |    0.0830 |  25.0620 |       4 |   0.3277 |
| synth_physics  | tabpfn         | heuristic   | -      | rmse     |    0.1108 |    0.0852 |  23.1440 |      20 |   0.3277 |
| wine           | tabpfn         | none        | -      | logloss  |    0.0085 |    0.0085 |   0.0000 |       1 |   0.0069 |

## T3. TabArena protocol vs the paper (per-dataset test error, P0 / P* vs TabFM / TabFM-Auto Opus 5)
| dataset                               |        P0 |        P* |   our_gain_% |     tabfm |   tabfm_auto_opus5 |   paper_gain_opus5_% | run                                                                 |
|:--------------------------------------|----------:|----------:|-------------:|----------:|-------------------:|---------------------:|:--------------------------------------------------------------------|
| hazelnut-spread-contaminant-detection |    0.0048 |    0.0047 |       1.3302 |    0.0023 |             0.0021 |               8.6957 | 20261002T155711Z_J2_heuristic_hazelnut-spread-contaminant-detection |
| qsar-biodeg                           |    0.0596 |    0.0603 |      -1.2004 |    0.0580 |             0.0582 |              -0.3448 | 20261002T155339Z_J2_heuristic_qsar-biodeg                           |
| diabetes                              |    0.1596 |    0.1623 |      -1.6644 |    0.1580 |             0.1471 |               6.8987 | 20261002T155211Z_J2_heuristic_diabetes                              |
| QSAR_fish_toxicity                    |    0.8547 |    0.8529 |       0.2184 |    0.8535 |             0.8483 |               0.6046 | 20261002T155220Z_J2_heuristic_QSAR_fish_toxicity                    |
| healthcare_insurance_expenses         | 4553.6089 | 4613.8992 |      -1.3240 | 4417.6000 |          4344.1000 |               1.6638 | 20261002T155549Z_J2_heuristic_healthcare_insurance_expenses         |
| maternal_health_risk                  |    0.3847 |    0.3952 |      -2.7479 |    0.3706 |             0.3648 |               1.5650 | 20261002T155220Z_J2_heuristic_maternal_health_risk                  |
| MIC                                   |    0.4318 |    0.4254 |       1.4831 |    0.4282 |             0.4178 |               2.4288 | 20261002T155711Z_J2_heuristic_MIC                                   |
| credit-g                              |    0.1979 |    0.1983 |      -0.1901 |    0.1944 |             0.1940 |               0.2058 | 20261002T155220Z_J2_heuristic_credit-g                              |
| blood-transfusion-service-center      |    0.2458 |    0.2453 |       0.2097 |    0.2441 |             0.2431 |               0.4097 | 20261002T155211Z_J2_heuristic_blood-transfusion-service-center      |
| website_phishing                      |    0.2144 |    0.2134 |       0.4375 |    0.2104 |             0.2067 |               1.7586 | 20261002T155610Z_J2_heuristic_website_phishing                      |
| Fitness_Club                          |    0.1787 |    0.1774 |       0.7010 |    0.1789 |             0.1787 |               0.1118 | 20261002T155643Z_J2_heuristic_Fitness_Club                          |
| concrete_compressive_strength         |    3.8268 |    3.8782 |      -1.3426 |    3.9666 |             3.8169 |               3.7740 | 20261002T155239Z_J2_heuristic_concrete_compressive_strength         |
| anneal                                |    0.0118 |    0.0108 |       8.4950 |    0.0125 |             0.0103 |              17.6000 | 20261002T155211Z_J2_heuristic_anneal                                |
| Marketing_Campaign                    |    0.0660 |    0.0659 |       0.1272 |    0.0732 |             0.0616 |              15.8470 | 20261002T155712Z_J2_heuristic_Marketing_Campaign                    |
| Is-this-a-good-customer               |    0.2474 |    0.2532 |      -2.3436 |    0.2466 |             0.2432 |               1.3788 | 20261002T155709Z_J2_heuristic_Is-this-a-good-customer               |
| Another-Dataset-on-used-Fiat-500      |  713.9073 |  719.0253 |      -0.7169 |  703.2700 |           693.4900 |               1.3906 | 20261002T155710Z_J2_heuristic_Another-Dataset-on-used-Fiat-500      |
| airfoil_self_noise                    |    0.9793 |    0.9161 |       6.4590 |    1.0734 |             0.9165 |              14.6171 | 20261002T155640Z_J2_heuristic_airfoil_self_noise                    |

## T4. Backbone transfer of discovered pipelines
|                                      |     P* |     P0 |   gain % |
|:-------------------------------------|-------:|-------:|---------:|
| ('synth_entities', 'hgb')            | 0.4377 | 0.4733 |   7.5093 |
| ('synth_entities', 'kumo-tabular-l') | 0.3496 | 0.3694 |   5.3387 |
| ('synth_entities', 'kumo-tabular-m') | 0.3687 | 0.3120 | -18.2011 |
| ('synth_entities', 'kumo-tabular-s') | 0.3242 | 0.4384 |  26.0379 |
| ('synth_entities', 'tabicl')         | 0.3812 | 0.3794 |  -0.4622 |
| ('synth_entities', 'tabpfn')         | 0.4126 | 0.4657 |  11.3975 |
| ('synth_entities', 'tabpfn-2.5')     | 0.4200 | 0.4862 |  13.6251 |
| ('synth_physics', 'hgb')             | 0.1813 | 0.3238 |  44.0186 |
| ('synth_physics', 'kumo-tabular-l')  | 0.0841 | 0.0832 |  -1.0959 |
| ('synth_physics', 'kumo-tabular-m')  | 0.0835 | 0.0842 |   0.8305 |
| ('synth_physics', 'kumo-tabular-s')  | 0.0871 | 0.0861 |  -1.0924 |
| ('synth_physics', 'tabicl')          | 0.0889 | 0.0928 |   4.1722 |
| ('synth_physics', 'tabpfn')          | 0.0883 | 0.1045 |  15.5013 |
| ('synth_physics', 'tabpfn-2.5')      | 0.1035 | 0.0930 | -11.2714 |

## T5. Retrieval benchmark (synthetic shop DB; mean/std over seeds)
|                                          |   seg P@10 mean |   seg P@10 std |   seg P@10 count |   seg kNN acc mean |   seg kNN acc std |   seg kNN acc count |   fut MAP@10 mean |   fut MAP@10 std |   fut MAP@10 count |
|:-----------------------------------------|----------------:|---------------:|-----------------:|-------------------:|------------------:|--------------------:|------------------:|-----------------:|-------------------:|
| ('synth_shop', 'agg')                    |           0.726 |          0.000 |            8.000 |              0.843 |             0.000 |               8.000 |             0.048 |            0.000 |              8.000 |
| ('synth_shop', 'gnn')                    |           0.879 |          0.006 |            8.000 |              0.903 |             0.005 |               8.000 |             0.047 |            0.001 |              8.000 |
| ('synth_shop', 'kumo_relational')        |           0.330 |          0.132 |            3.000 |              0.422 |             0.155 |               3.000 |             0.036 |            0.009 |              3.000 |
| ('synth_shop', 'kumo_relational_kmeans') |           0.615 |          0.051 |            3.000 |              0.695 |             0.037 |               3.000 |             0.045 |            0.004 |              3.000 |
| ('synth_shop', 'openrfm')                |           0.249 |          0.000 |            8.000 |              0.306 |             0.000 |               8.000 |             0.032 |            0.000 |              8.000 |
| ('synth_shop', 'openrfm_ctx16')          |           0.237 |        nan     |            1.000 |              0.291 |           nan     |               1.000 |             0.034 |          nan     |              1.000 |
| ('synth_shop', 'openrfm_random')         |           0.168 |          0.000 |            4.000 |              0.169 |             0.000 |               4.000 |             0.013 |            0.000 |              4.000 |
| ('synth_shop', 'row')                    |           0.161 |          0.000 |            8.000 |              0.146 |             0.000 |               8.000 |             0.022 |            0.000 |              8.000 |
| ('synth_shop', 'tabpfn_agg')             |           0.547 |          0.315 |            7.000 |              0.617 |             0.290 |               7.000 |             0.038 |            0.012 |              7.000 |
| ('synth_shop', 'tabpfn_agg_churn')       |           0.182 |        nan     |            1.000 |              0.173 |           nan     |               1.000 |             0.022 |          nan     |              1.000 |
| ('synth_shop', 'tabpfn_agg_feat4')       |           0.680 |          0.030 |            3.000 |              0.787 |             0.028 |               3.000 |             0.042 |            0.001 |              3.000 |
| ('synth_shop', 'tabpfn_agg_kmeans')      |           0.823 |          0.018 |            6.000 |              0.874 |             0.012 |               6.000 |             0.048 |            0.001 |              6.000 |
| ('synth_shop', 'tabpfn_agg_rand4')       |           0.525 |          0.325 |            3.000 |              0.617 |             0.314 |               3.000 |             0.037 |            0.012 |              3.000 |
| ('synth_shop', 'tabpfn_row')             |           0.171 |          0.005 |            4.000 |              0.174 |             0.012 |               4.000 |             0.021 |            0.001 |              4.000 |

## T6. RelBench rel-hm user-item-purchase (official evaluator)
_no RelBench runs yet_
