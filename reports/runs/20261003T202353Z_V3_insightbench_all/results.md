# P3: InsightBench (100 tables), analyst agent ablation

D = LLM + SQL; E = + DEEP tools (frozen tabular models); F = + GLiClass text labels + verified insight ledger. g_eval = open-judge match of each planted insight (0-1, mean over gold insights); rouge1 as in the benchmark.

| config   |   g_eval |   rouge1 |   tables |   insights |   rejected |   llm_calls |   seconds |   errors |   vs_D |   vs_D_se |
|:---------|---------:|---------:|---------:|-----------:|-----------:|------------:|----------:|---------:|-------:|----------:|
| F        |   0.2301 |   0.2009 |      100 |       3.16 |       2.53 |       20.27 |   106.574 |        0 |    nan |       nan |
