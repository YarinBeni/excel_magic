# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |    vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|--------:|----------:|
| piD      |   0.2727 |   0.194  |      100 |       6.77 |       0    |           0 |    26.786 |        0 |  0      |    0      |
| piE      |   0.2687 |   0.1911 |      100 |       6.73 |       0    |           0 |    27.421 |        0 | -0.004  |    0.0264 |
| piF      |   0.2102 |   0.1607 |      100 |       5.2  |       5.76 |           0 |    35.526 |        0 | -0.0625 |    0.03   |
