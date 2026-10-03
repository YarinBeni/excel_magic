# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |    vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|--------:|----------:|
| D        |   0.4333 |   0.2569 |        5 |        6.4 |          0 |        16.4 |     11.4  |        0 |  0      |    0      |
| E        |   0.4065 |   0.2557 |        5 |        5.8 |          0 |        18.8 |     19.36 |        0 | -0.0269 |    0.0557 |
| F        |   0.3648 |   0.2463 |        5 |        2.4 |          0 |        15.8 |     20.5  |        0 | -0.0685 |    0.0404 |
