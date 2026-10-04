# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |    vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|--------:|----------:|
| D        |   0.2781 |   0.2371 |      100 |       5.61 |          0 |       15.16 |    13.482 |        0 |  0      |    0      |
| Q        |   0.2342 |   0.2295 |      100 |       5.77 |          0 |       16.64 |    20.053 |        0 | -0.0439 |    0.0189 |
| Q0       |   0.2176 |   0.2282 |      100 |       5.84 |          0 |       16.11 |    17.943 |        0 | -0.0605 |    0.016  |
