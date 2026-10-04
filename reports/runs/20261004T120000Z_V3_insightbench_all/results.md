# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |    vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|--------:|----------:|
| D        |   0.3149 |   0.2394 |      100 |       5.67 |          0 |       15.06 |    12.868 |        0 |  0      |    0      |
| P        |   0.2982 |   0.2245 |      100 |       5.83 |          0 |       15.03 |    14.062 |        0 | -0.0167 |    0.0198 |
| PC       |   0.2803 |   0.2279 |      100 |       5.97 |          0 |       16.15 |    15.458 |        0 | -0.0346 |    0.0199 |
