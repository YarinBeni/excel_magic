# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |    vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|--------:|----------:|
| D        |   0.3043 |   0.2447 |      100 |       5.62 |       0    |       14.88 |    12.903 |        0 |  0      |    0      |
| E        |   0.2664 |   0.2352 |      100 |       4.85 |       0    |       17.58 |    87.745 |        0 | -0.0379 |    0.0177 |
| F        |   0.2234 |   0.1997 |      100 |       3.07 |       2.47 |       20.46 |    97.981 |        0 | -0.0809 |    0.0215 |
