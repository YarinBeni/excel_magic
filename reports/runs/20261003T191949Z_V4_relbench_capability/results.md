# P3: prediction questions over RelBench databases, agent configs

D = LLM + SQL; E = + Kumo Tabular-L fitted on LLM-built features; R = + Kumo Relational (graph layer probe); ER = both. Official test AUROC, mean over repeats (unsubmitted = missing).

| task              |      D |      E |     ER |      R |
|:------------------|-------:|-------:|-------:|-------:|
| rel-f1/driver-dnf | 0.4187 | 0.7109 | 0.7482 | 0.7482 |

Share of episodes that submitted:

| task              |   D |   E |   ER |   R |
|:------------------|----:|----:|-----:|----:|
| rel-f1/driver-dnf |   1 |   1 |    1 |   1 |

Mean over tasks: {'D': 0.4187, 'E': 0.7109, 'ER': 0.7482, 'R': 0.7482}
