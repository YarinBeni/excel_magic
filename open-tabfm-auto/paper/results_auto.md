# Auto-generated results tables
_from 158 runs under runs, /home/user/excel_magic/reports/runs, /home/user/frozen-embeddings-retrieval/runs_

## T1. Frozen backbones with the identity pipeline (3-fold CV, lower is better)
|                               |     hgb |   kumo-tabular-l |   kumo-tabular-m |   kumo-tabular-s |   tabicl |   tabpfn |   tabpfn-2.5 |
|:------------------------------|--------:|-----------------:|-----------------:|-----------------:|---------:|---------:|-------------:|
| ('breast_cancer', '1-auroc')  |  0.0059 |           0.0033 |           0.0038 |           0.0039 |   0.0042 |   0.0042 |       0.0039 |
| ('diabetes', 'rmse')          | 60.5173 |          54.5179 |          54.1992 |          54.5932 |  53.7263 |  54.2821 |      54.1759 |
| ('synth_entities', '1-auroc') |  0.4873 |           0.4319 |           0.4213 |           0.4424 |   0.4439 |   0.5010 |       0.4492 |
| ('synth_physics', 'rmse')     |  0.3539 |           0.0838 |           0.0854 |           0.0868 |   0.0951 |   0.0920 |       0.0961 |
| ('wine', 'logloss')           |  0.0965 |           0.0175 |           0.0186 |           0.0294 |   0.0158 |   0.0356 |       0.0361 |

## T2. Pipeline search: harness x LLM x backbone (held-out test error)
| dataset        | backbone       | harness     | llm                               | metric   |   P0 test |   P* test |   gain % |   evals |   HistGB |
|:---------------|:---------------|:------------|:----------------------------------|:---------|----------:|----------:|---------:|--------:|---------:|
| breast_cancer  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0045 |    0.0050 |  -9.6770 |       1 |   0.0108 |
| breast_cancer  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0051 |    0.0072 | -40.0000 |       9 |   0.0108 |
| breast_cancer  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0045 |    0.0053 | -16.1290 |       6 |   0.0108 |
| breast_cancer  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0051 |    0.0060 | -17.1430 |      12 |   0.0108 |
| breast_cancer  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0054 |    0.0057 |  -5.4050 |       9 |   0.0108 |
| breast_cancer  | kumo-tabular-s | heuristic   | -                                 | 1-auroc  |    0.0050 |    0.0050 |   0.0000 |      16 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | Qwen/Qwen3-32B                    | 1-auroc  |    0.0051 |    0.0053 |  -2.8570 |       6 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0051 |    0.0053 |  -2.8570 |      16 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0044 |    0.0069 | -56.6670 |       9 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.0050 |    0.0061 | -23.5290 |      12 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | openai/gpt-oss-20b                | 1-auroc  |    0.0055 |    0.0051 |   7.8950 |       6 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | openai/gpt-oss-20b                | 1-auroc  |    0.0055 |    0.0060 |  -7.8950 |      10 |   0.0108 |
| breast_cancer  | kumo-tabular-s | openai      | zai-org/GLM-4.5-Air-FP8           | 1-auroc  |    0.0053 |    0.0073 | -38.8890 |       9 |   0.0108 |
| breast_cancer  | tabpfn         | heuristic   | -                                 | 1-auroc  |    0.0061 |    0.0061 |   0.0000 |      20 |   0.0108 |
| synth_entities | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4688 |    0.4585 |   2.1950 |       1 |   0.4733 |
| synth_entities | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.3842 |    0.2239 |  41.7220 |       9 |   0.4733 |
| synth_entities | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4592 |    0.3422 |  25.4840 |       9 |   0.4733 |
| synth_entities | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4515 |    0.2242 |  50.3370 |       6 |   0.4733 |
| synth_entities | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4517 |    0.3248 |  28.0990 |       9 |   0.4733 |
| synth_entities | kumo-tabular-s | heuristic   | -                                 | 1-auroc  |    0.4623 |    0.3281 |  29.0230 |      20 |   0.4733 |
| synth_entities | kumo-tabular-s | heuristic   | -                                 | 1-auroc  |    0.4412 |    0.3327 |  24.5890 |      20 |   0.4733 |
| synth_entities | kumo-tabular-s | heuristic   | -                                 | 1-auroc  |    0.3740 |    0.3429 |   8.3150 |      16 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | Qwen/Qwen3-32B                    | 1-auroc  |    0.4546 |    0.4718 |  -3.7890 |      16 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4387 |    0.4437 |  -1.1460 |       9 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4545 |    0.4499 |   1.0030 |      12 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | 1-auroc  |    0.4440 |    0.4658 |  -4.9150 |      10 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | openai/gpt-oss-20b                | 1-auroc  |    0.4307 |    0.4726 |  -9.7370 |       6 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | openai/gpt-oss-20b                | 1-auroc  |    0.3803 |    0.4144 |  -8.9450 |       9 |   0.4733 |
| synth_entities | kumo-tabular-s | openai      | zai-org/GLM-4.5-Air-FP8           | 1-auroc  |    0.3972 |    0.2240 |  43.6130 |      16 |   0.4733 |
| synth_entities | tabpfn         | heuristic   | -                                 | 1-auroc  |    0.4676 |    0.4266 |   8.7750 |      20 |   0.4585 |
| synth_physics  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0877 |    0.0859 |   2.1040 |       1 |   0.3238 |
| synth_physics  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0861 |    0.0880 |  -2.1770 |       9 |   0.3238 |
| synth_physics  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0858 |    0.0849 |   1.0080 |       9 |   0.3238 |
| synth_physics  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0880 |    0.0888 |  -0.9190 |      10 |   0.3238 |
| synth_physics  | kumo-tabular-s | cli         | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0863 |    0.0864 |  -0.0550 |      10 |   0.3238 |
| synth_physics  | kumo-tabular-s | heuristic   | -                                 | rmse     |    0.0876 |    0.0872 |   0.5530 |      20 |   0.3238 |
| synth_physics  | kumo-tabular-s | heuristic   | -                                 | rmse     |    0.0872 |    0.0872 |   0.0350 |      20 |   0.3238 |
| synth_physics  | kumo-tabular-s | heuristic   | -                                 | rmse     |    0.0866 |    0.0874 |  -0.9510 |      16 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | Qwen/Qwen3-32B                    | rmse     |    0.0862 |    0.0863 |  -0.1880 |       1 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0869 |    0.0857 |   1.3240 |       8 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0865 |    0.0840 |   2.9280 |      12 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | Qwen/Qwen3-Coder-30B-A3B-Instruct | rmse     |    0.0860 |    0.0866 |  -0.7120 |       8 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | openai/gpt-oss-20b                | rmse     |    0.0864 |    0.0875 |  -1.2590 |       2 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | openai/gpt-oss-20b                | rmse     |    0.0869 |    0.0834 |   4.0730 |       6 |   0.3238 |
| synth_physics  | kumo-tabular-s | openai      | zai-org/GLM-4.5-Air-FP8           | rmse     |    0.0884 |    0.0839 |   5.0630 |      16 |   0.3238 |
| synth_physics  | tabpfn         | claude-code | sonnet                            | rmse     |    0.1108 |    0.0830 |  25.0620 |       4 |   0.3277 |
| synth_physics  | tabpfn         | heuristic   | -                                 | rmse     |    0.1108 |    0.0852 |  23.1440 |      20 |   0.3277 |
| wine           | tabpfn         | none        | -                                 | logloss  |    0.0085 |    0.0085 |   0.0000 |       1 |   0.0069 |

## T3. TabArena protocol vs the paper (per-dataset test error, P0 / P* vs TabFM / TabFM-Auto Opus 5)
| dataset                               |        P0 |        P* |   our_gain_% |     tabfm |   tabfm_auto_opus5 |   paper_gain_opus5_% | run                                                                          |
|:--------------------------------------|----------:|----------:|-------------:|----------:|-------------------:|---------------------:|:-----------------------------------------------------------------------------|
| airfoil_self_noise                    |    0.8222 |    0.7953 |       3.2720 |    1.0734 |             0.9165 |              14.6171 | 20261002T204156Z_J13_rich_qwen3coder_tabarena                                |
| blood-transfusion-service-center      |    0.2650 |    0.2639 |       0.4303 |    0.2441 |             0.2431 |               0.4097 | 20261002T204156Z_J13_rich_qwen3coder_tabarena                                |
| credit-g                              |    0.2071 |    0.2043 |       1.3617 |    0.1944 |             0.1940 |               0.2058 | 20261002T204156Z_J13_rich_qwen3coder_tabarena                                |
| credit-g                              |    0.1925 |    0.1957 |      -1.6254 |    0.1944 |             0.1940 |               0.2058 | 20261002T204301Z_J9_heuristic_kumoL_credit-g                                 |
| hazelnut-spread-contaminant-detection |    0.0048 |    0.0047 |       1.3302 |    0.0023 |             0.0021 |               8.6957 | 20261002T155711Z_J2_heuristic_hazelnut-spread-contaminant-detection          |
| Another-Dataset-on-used-Fiat-500      |  714.4997 |  717.2734 |      -0.3882 |  703.2700 |           693.4900 |               1.3906 | 20261002T195318Z_J8_pi_qwen3coder_Another-Dataset-on-used-Fiat-500           |
| qsar-biodeg                           |    0.0596 |    0.0603 |      -1.2004 |    0.0580 |             0.0582 |              -0.3448 | 20261002T155339Z_J2_heuristic_qsar-biodeg                                    |
| blood-transfusion-service-center      |    0.2451 |    0.2462 |      -0.4498 |    0.2441 |             0.2431 |               0.4097 | 20261002T194110Z_J8_openai_glm45air_blood-transfusion-service-center         |
| concrete_compressive_strength         |    3.8471 |    3.7847 |       1.6205 |    3.9666 |             3.8169 |               3.7740 | 20261002T201544Z_J8_openai_glm45air_concrete_compressive_strength            |
| diabetes                              |    0.1596 |    0.1623 |      -1.6644 |    0.1580 |             0.1471 |               6.8987 | 20261002T155211Z_J2_heuristic_diabetes                                       |
| airfoil_self_noise                    |    0.9786 |    0.9422 |       3.7172 |    1.0734 |             0.9165 |              14.6171 | 20261002T194956Z_J8_pi_qwen3coder_airfoil_self_noise                         |
| healthcare_insurance_expenses         | 4555.6918 | 4539.1575 |       0.3629 | 4417.6000 |          4344.1000 |               1.6638 | 20261002T202730Z_J8_openai_glm45air_healthcare_insurance_expenses            |
| QSAR_fish_toxicity                    |    0.8549 |    0.8531 |       0.2140 |    0.8535 |             0.8483 |               0.6046 | 20261002T204302Z_J11_heuristic_cv3_QSAR_fish_toxicity                        |
| diabetes                              |    0.1596 |    0.1597 |      -0.0493 |    0.1580 |             0.1471 |               6.8987 | 20261002T203810Z_J10_pi_qwen3coder_b64_diabetes                              |
| diabetes                              |    0.1611 |    0.1616 |      -0.2854 |    0.1580 |             0.1471 |               6.8987 | 20261002T202622Z_J9_heuristic_tabiclv2_diabetes                              |
| qsar-biodeg                           |    0.0596 |    0.0595 |       0.1055 |    0.0580 |             0.0582 |              -0.3448 | 20261002T193940Z_J8_pi_qwen3coder_qsar-biodeg                                |
| blood-transfusion-service-center      |    0.2446 |    0.2467 |      -0.8823 |    0.2441 |             0.2431 |               0.4097 | 20261002T202255Z_J9_heuristic_tabiclv2_blood-transfusion-service-center      |
| anneal                                |    0.0116 |    0.0117 |      -0.8113 |    0.0125 |             0.0103 |              17.6000 | 20261002T203110Z_J10_openai_glm45air_b64_anneal                              |
| anneal                                |    0.0116 |    0.0122 |      -5.2112 |    0.0125 |             0.0103 |              17.6000 | 20261002T192430Z_J8_pi_qwen3coder_anneal                                     |
| credit-g                              |    0.2031 |    0.2060 |      -1.4258 |    0.1944 |             0.1940 |               0.2058 | 20261002T203805Z_J9_heuristic_tabiclv2_credit-g                              |
| QSAR_fish_toxicity                    |    0.8565 |    0.8529 |       0.4250 |    0.8535 |             0.8483 |               0.6046 | 20261002T203813Z_J9_heuristic_kumoL_QSAR_fish_toxicity                       |
| QSAR_fish_toxicity                    |    0.8547 |    0.8529 |       0.2184 |    0.8535 |             0.8483 |               0.6046 | 20261002T155220Z_J2_heuristic_QSAR_fish_toxicity                             |
| healthcare_insurance_expenses         | 4553.6089 | 4613.8992 |      -1.3240 | 4417.6000 |          4344.1000 |               1.6638 | 20261002T155549Z_J2_heuristic_healthcare_insurance_expenses                  |
| concrete_compressive_strength         |    3.8327 |    3.8148 |       0.4664 |    3.9666 |             3.8169 |               3.7740 | 20261002T193646Z_J8_pi_qwen3coder_concrete_compressive_strength              |
| anneal                                |    0.0179 |    0.0274 |     -52.4885 |    0.0125 |             0.0103 |              17.6000 | 20261002T202959Z_J9_heuristic_tabiclv2_anneal                                |
| maternal_health_risk                  |    0.3847 |    0.3952 |      -2.7479 |    0.3706 |             0.3648 |               1.5650 | 20261002T155220Z_J2_heuristic_maternal_health_risk                           |
| Is-this-a-good-customer               |    0.2475 |    0.2475 |      -0.0213 |    0.2466 |             0.2432 |               1.3788 | 20261002T195952Z_J8_pi_qwen3coder_Is-this-a-good-customer                    |
| website_phishing                      |    0.2146 |    0.2149 |      -0.1367 |    0.2104 |             0.2067 |               1.7586 | 20261002T194521Z_J8_pi_qwen3coder_website_phishing                           |
| MIC                                   |    0.4318 |    0.4254 |       1.4831 |    0.4282 |             0.4178 |               2.4288 | 20261002T155711Z_J2_heuristic_MIC                                            |
| blood-transfusion-service-center      |    0.2475 |    0.2495 |      -0.8143 |    0.2441 |             0.2431 |               0.4097 | 20261002T202026Z_J9_heuristic_kumoL_blood-transfusion-service-center         |
| anneal                                |    0.0110 |    0.0120 |      -8.8442 |    0.0125 |             0.0103 |              17.6000 | 20261002T203215Z_J9_heuristic_kumoL_anneal                                   |
| blood-transfusion-service-center      |    0.2454 |    0.2456 |      -0.0836 |    0.2441 |             0.2431 |               0.4097 | 20261002T204556Z_J11_pi_qwen3coder_cv3_blood-transfusion-service-center      |
| airfoil_self_noise                    |    0.8264 |    0.8337 |      -0.8860 |    1.0734 |             0.9165 |              14.6171 | 20261002T175823Z_J5_GLM-4.5-Air-FP8_tabarena                                 |
| blood-transfusion-service-center      |    0.2639 |    0.2602 |       1.4290 |    0.2441 |             0.2431 |               0.4097 | 20261002T175823Z_J5_GLM-4.5-Air-FP8_tabarena                                 |
| credit-g                              |    0.2069 |    0.2039 |       1.4460 |    0.1944 |             0.1940 |               0.2058 | 20261002T175823Z_J5_GLM-4.5-Air-FP8_tabarena                                 |
| anneal                                |    0.0118 |    0.0122 |      -2.7376 |    0.0125 |             0.0103 |              17.6000 | 20261002T195220Z_J8_openai_glm45air_anneal                                   |
| credit-g                              |    0.1979 |    0.1983 |      -0.1901 |    0.1944 |             0.1940 |               0.2058 | 20261002T155220Z_J2_heuristic_credit-g                                       |
| blood-transfusion-service-center      |    0.2458 |    0.2453 |       0.2097 |    0.2441 |             0.2431 |               0.4097 | 20261002T155211Z_J2_heuristic_blood-transfusion-service-center               |
| hazelnut-spread-contaminant-detection |    0.0049 |    0.0051 |      -3.9991 |    0.0023 |             0.0021 |               8.6957 | 20261002T203505Z_J10_pi_qwen3coder_b64_hazelnut-spread-contaminant-detection |
| website_phishing                      |    0.2141 |    0.2154 |      -0.5939 |    0.2104 |             0.2067 |               1.7586 | 20261002T203348Z_J8_openai_glm45air_website_phishing                         |
| website_phishing                      |    0.2144 |    0.2134 |       0.4375 |    0.2104 |             0.2067 |               1.7586 | 20261002T155610Z_J2_heuristic_website_phishing                               |
| Fitness_Club                          |    0.1787 |    0.1774 |       0.7010 |    0.1789 |             0.1787 |               0.1118 | 20261002T155643Z_J2_heuristic_Fitness_Club                                   |
| concrete_compressive_strength         |    3.8268 |    3.8782 |      -1.3426 |    3.9666 |             3.8169 |               3.7740 | 20261002T155239Z_J2_heuristic_concrete_compressive_strength                  |
| blood-transfusion-service-center      |    0.2451 |    0.2454 |      -0.1416 |    0.2441 |             0.2431 |               0.4097 | 20261002T191905Z_J8_pi_qwen3coder_blood-transfusion-service-center           |
| anneal                                |    0.0118 |    0.0108 |       8.4950 |    0.0125 |             0.0103 |              17.6000 | 20261002T155211Z_J2_heuristic_anneal                                         |
| maternal_health_risk                  |    0.3986 |    0.3985 |       0.0482 |    0.3706 |             0.3648 |               1.5650 | 20261002T204144Z_J9_heuristic_tabiclv2_maternal_health_risk                  |
| diabetes                              |    0.1597 |    0.1578 |       1.2164 |    0.1580 |             0.1471 |               6.8987 | 20261002T203818Z_J11_heuristic_cv3_diabetes                                  |
| hazelnut-spread-contaminant-detection |    0.0048 |    0.0048 |       0.2471 |    0.0023 |             0.0021 |               8.6957 | 20261002T200518Z_J8_pi_qwen3coder_hazelnut-spread-contaminant-detection      |
| airfoil_self_noise                    |    0.8318 |    0.8261 |       0.6887 |    1.0734 |             0.9165 |              14.6171 | 20261002T182208Z_J5_Qwen3-32B_tabarena                                       |
| blood-transfusion-service-center      |    0.2656 |    0.2639 |       0.6605 |    0.2441 |             0.2431 |               0.4097 | 20261002T182208Z_J5_Qwen3-32B_tabarena                                       |
| credit-g                              |    0.2062 |    0.2055 |       0.3523 |    0.1944 |             0.1940 |               0.2058 | 20261002T182208Z_J5_Qwen3-32B_tabarena                                       |
| Marketing_Campaign                    |    0.0660 |    0.0659 |       0.1272 |    0.0732 |             0.0616 |              15.8470 | 20261002T155712Z_J2_heuristic_Marketing_Campaign                             |
| airfoil_self_noise                    |    0.9749 |    0.9281 |       4.7997 |    1.0734 |             0.9165 |              14.6171 | 20261002T203824Z_J10_heuristic_b64_airfoil_self_noise                        |
| airfoil_self_noise                    |    0.9775 |    0.9530 |       2.5067 |    1.0734 |             0.9165 |              14.6171 | 20261002T203207Z_J10_pi_qwen3coder_b64_airfoil_self_noise                    |
| healthcare_insurance_expenses         | 4559.0442 | 4521.9604 |       0.8134 | 4417.6000 |          4344.1000 |               1.6638 | 20261002T194237Z_J8_pi_qwen3coder_healthcare_insurance_expenses              |
| QSAR_fish_toxicity                    |    0.8584 |    0.8585 |      -0.0056 |    0.8535 |             0.8483 |               0.6046 | 20261002T203422Z_J9_heuristic_tabiclv2_QSAR_fish_toxicity                    |
| Is-this-a-good-customer               |    0.2474 |    0.2532 |      -2.3436 |    0.2466 |             0.2432 |               1.3788 | 20261002T155709Z_J2_heuristic_Is-this-a-good-customer                        |
| Another-Dataset-on-used-Fiat-500      |  713.9073 |  719.0253 |      -0.7169 |  703.2700 |           693.4900 |               1.3906 | 20261002T155710Z_J2_heuristic_Another-Dataset-on-used-Fiat-500               |
| Marketing_Campaign                    |    0.0661 |    0.0755 |     -14.2093 |    0.0732 |             0.0616 |              15.8470 | 20261002T200211Z_J8_pi_qwen3coder_Marketing_Campaign                         |
| Marketing_Campaign                    |    0.0658 |    0.0759 |     -15.2510 |    0.0732 |             0.0616 |              15.8470 | 20261002T202911Z_J10_pi_qwen3coder_b64_Marketing_Campaign                    |
| Marketing_Campaign                    |    0.0657 |    0.0662 |      -0.7262 |    0.0732 |             0.0616 |              15.8470 | 20261002T203313Z_J10_heuristic_b64_Marketing_Campaign                        |
| diabetes                              |    0.1595 |    0.1582 |       0.8293 |    0.1580 |             0.1471 |               6.8987 | 20261002T192134Z_J8_pi_qwen3coder_diabetes                                   |
| QSAR_fish_toxicity                    |    0.8549 |    0.8546 |       0.0334 |    0.8535 |             0.8483 |               0.6046 | 20261002T195829Z_J8_openai_glm45air_QSAR_fish_toxicity                       |
| MIC                                   |    0.4324 |    0.4329 |      -0.1076 |    0.4282 |             0.4178 |               2.4288 | 20261002T195641Z_J8_pi_qwen3coder_MIC                                        |
| airfoil_self_noise                    |    0.9793 |    0.9161 |       6.4590 |    1.0734 |             0.9165 |              14.6171 | 20261002T155640Z_J2_heuristic_airfoil_self_noise                             |
| Marketing_Campaign                    |    0.0659 |    0.0973 |     -47.6731 |    0.0732 |             0.0616 |              15.8470 | 20261002T203854Z_J10_openai_glm45air_b64_Marketing_Campaign                  |
| qsar-biodeg                           |    0.0596 |    0.0598 |      -0.2938 |    0.0580 |             0.0582 |              -0.3448 | 20261002T202129Z_J8_openai_glm45air_qsar-biodeg                              |
| QSAR_fish_toxicity                    |    0.8546 |    0.8551 |      -0.0646 |    0.8535 |             0.8483 |               0.6046 | 20261002T192744Z_J8_pi_qwen3coder_QSAR_fish_toxicity                         |
| Fitness_Club                          |    0.1786 |    0.1783 |       0.1872 |    0.1789 |             0.1787 |               0.1118 | 20261002T203852Z_J8_openai_glm45air_Fitness_Club                             |
| credit-g                              |    0.1980 |    0.2024 |      -2.2303 |    0.1944 |             0.1940 |               0.2058 | 20261002T193116Z_J8_pi_qwen3coder_credit-g                                   |
| anneal                                |    0.0115 |    0.0120 |      -4.5678 |    0.0125 |             0.0103 |              17.6000 | 20261002T202510Z_J10_pi_qwen3coder_b64_anneal                                |
| blood-transfusion-service-center      |    0.2455 |    0.2469 |      -0.5656 |    0.2441 |             0.2431 |               0.4097 | 20261002T203336Z_J11_heuristic_cv3_blood-transfusion-service-center          |
| airfoil_self_noise                    |    0.8277 |    0.8088 |       2.2844 |    1.0734 |             0.9165 |              14.6171 | 20261002T173650Z_J5_Qwen3-Coder-30B-A3B-Instruct_tabarena                    |
| blood-transfusion-service-center      |    0.2631 |    0.2655 |      -0.9336 |    0.2441 |             0.2431 |               0.4097 | 20261002T173650Z_J5_Qwen3-Coder-30B-A3B-Instruct_tabarena                    |
| credit-g                              |    0.2054 |    0.2062 |      -0.3537 |    0.1944 |             0.1940 |               0.2058 | 20261002T173650Z_J5_Qwen3-Coder-30B-A3B-Instruct_tabarena                    |
| diabetes                              |    0.1564 |    0.1592 |      -1.7509 |    0.1580 |             0.1471 |               6.8987 | 20261002T202533Z_J9_heuristic_kumoL_diabetes                                 |
| concrete_compressive_strength         |    3.9654 |    4.0204 |      -1.3860 |    3.9666 |             3.8169 |               3.7740 | 20261002T204512Z_J9_heuristic_tabiclv2_concrete_compressive_strength         |
| credit-g                              |    0.1976 |    0.2008 |      -1.6505 |    0.1944 |             0.1940 |               0.2058 | 20261002T200408Z_J8_openai_glm45air_credit-g                                 |
| airfoil_self_noise                    |    0.8173 |    0.8007 |       2.0251 |    1.0734 |             0.9165 |              14.6171 | 20261002T181125Z_J5_gpt-oss-20b_tabarena                                     |
| blood-transfusion-service-center      |    0.2627 |    0.2657 |      -1.1352 |    0.2441 |             0.2431 |               0.4097 | 20261002T181125Z_J5_gpt-oss-20b_tabarena                                     |
| credit-g                              |    0.2071 |    0.2076 |      -0.2476 |    0.1944 |             0.1940 |               0.2058 | 20261002T181125Z_J5_gpt-oss-20b_tabarena                                     |
| anneal                                |    0.0118 |    0.0112 |       5.6938 |    0.0125 |             0.0103 |              17.6000 | 20261002T202358Z_J10_heuristic_b64_anneal                                    |
| maternal_health_risk                  |    0.3843 |    0.3835 |       0.2021 |    0.3706 |             0.3648 |               1.5650 | 20261002T193345Z_J8_pi_qwen3coder_maternal_health_risk                       |
| diabetes                              |    0.1595 |    0.1574 |       1.3163 |    0.1580 |             0.1471 |               6.8987 | 20261002T194551Z_J8_openai_glm45air_diabetes                                 |
| maternal_health_risk                  |    0.3845 |    0.4377 |     -13.8281 |    0.3706 |             0.3648 |               1.5650 | 20261002T200854Z_J8_openai_glm45air_maternal_health_risk                     |
| Fitness_Club                          |    0.1786 |    0.1783 |       0.1843 |    0.1789 |             0.1787 |               0.1118 | 20261002T194749Z_J8_pi_qwen3coder_Fitness_Club                               |
| airfoil_self_noise                    |    0.8209 |    0.8282 |      -0.8832 |    1.0734 |             0.9165 |              14.6171 | 20261002T173553Z_J3_tabarena_openai                                          |
| blood-transfusion-service-center      |    0.2647 |    0.2582 |       2.4520 |    0.2441 |             0.2431 |               0.4097 | 20261002T173553Z_J3_tabarena_openai                                          |
| credit-g                              |    0.2056 |    0.2062 |      -0.2911 |    0.1944 |             0.1940 |               0.2058 | 20261002T173553Z_J3_tabarena_openai                                          |

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
| method                              |   test |   val |
|:------------------------------------|-------:|------:|
| GlobalPopularity                    |  0.292 | 0.342 |
| ItemKNN                             |  1.073 | 1.035 |
| Past+ItemKNN                        |  2.197 | 1.901 |
| Past+kNN-CF[agg]                    |  2.184 | 1.889 |
| Past+kNN-CF[purchase_matrix]        |  2.216 | 1.919 |
| Past+kNN-CF[row]                    |  2.160 | 1.857 |
| Past+kNN-CF[svd]                    |  2.195 | 1.894 |
| Past+kNN-CF[svd_agg]                |  2.185 | 1.888 |
| Past+kNN-CF[tabpfn_agg_kmeans]      |  2.187 | 1.889 |
| Past+kNN-CF[tabpfn_agg_random]      |  2.177 | 1.886 |
| Past+kNN-CF[tabpfn_row_kmeans]      |  2.161 | 1.857 |
| Past+kNN-CF[tabpfn_svd_kmeans]      |  2.180 | 1.886 |
| Past+kNN-CF[tabpfn_svd_random]      |  2.182 | 1.885 |
| PastVisit                           |  2.199 | 1.904 |
| PastVisit+kNN-CF[agg]               |  2.191 | 1.897 |
| PastVisit+kNN-CF[row]               |  2.191 | 1.897 |
| PastVisit+kNN-CF[tabpfn_agg_kmeans] |  2.191 | 1.897 |
| PastVisit+kNN-CF[tabpfn_agg_random] |  2.191 | 1.897 |
| PastVisit+kNN-CF[tabpfn_row_kmeans] |  2.191 | 1.897 |
| kNN-CF[agg]                         |  0.259 | 0.258 |
| kNN-CF[purchase_matrix]             |  1.200 | 1.162 |
| kNN-CF[row]                         |  0.144 | 0.136 |
| kNN-CF[svd]                         |  0.833 | 0.822 |
| kNN-CF[svd_agg]                     |  0.654 | 0.661 |
| kNN-CF[tabpfn_agg_kmeans]           |  0.227 | 0.248 |
| kNN-CF[tabpfn_agg_random]           |  0.170 | 0.207 |
| kNN-CF[tabpfn_row_kmeans]           |  0.140 | 0.135 |
| kNN-CF[tabpfn_svd_kmeans]           |  0.419 | 0.401 |
| kNN-CF[tabpfn_svd_random]           |  0.475 | 0.514 |

## T7. RelBench rel-hm user-churn, entity-level kNN probe (AUROC, random customer folds)
| method                 |   AUROC |
|:-----------------------|--------:|
| Supervised[hgb on agg] |  0.6725 |
| kNN[agg]               |  0.6530 |
| kNN[tabpfn_agg_kmeans] |  0.6484 |
| kNN[tabpfn_svd_kmeans] |  0.6450 |
| kNN[tabpfn_agg_random] |  0.6425 |
| kNN[svd_agg]           |  0.6375 |
| kNN[svd]               |  0.5901 |
| kNN[row]               |  0.5198 |
| MajorityPrior          |  0.5000 |
