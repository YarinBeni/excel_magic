# P2: SQL harness ablation on BIRD Arcwise-Plat

A = LLM alone (greedy, full schema); B = + GLiClass schema linking; C = + GLiClass triage and one revision; D = + 8 candidates picked by a cheap verifier stack (fitted on the other half of the questions); F = D + LLM judge on the uncertain band; SC = 8 candidates, majority result (no small models).

| config   |   accuracy |   n |   llm_calls |   judge_calls |   seconds |   schema_chars |   tokens |    vs_A |   vs_A_se |
|:---------|-----------:|----:|------------:|--------------:|----------:|---------------:|---------:|--------:|----------:|
| A        |     0.6928 | 498 |      1      |         0     |    1.4629 |        4994.1  |  73.4357 |  0      |    0      |
| B        |     0.6667 | 498 |      1      |         0     |    2.57   |        2253.77 |  81.257  | -0.0261 |    0.0159 |
| C        |     0.6767 | 498 |      1.0301 |         0     |    2.6923 |        2253.77 |  79.3936 | -0.0161 |    0.0161 |
| D        |     0.7048 | 498 |      2.2711 |         0     |    9.264  |        2253.77 | 656.098  |  0.012  |    0.0158 |
| F        |     0.7028 | 498 |      2.2631 |         1.508 |    9.0407 |        2253.77 | 651.018  |  0.01   |    0.0154 |
| SC       |     0.7048 | 498 |      2      |         0     |    4.0253 |        4994.1  | 588.219  |  0.012  |    0.009  |
