# P3: prediction questions over RelBench databases, agent configs

D = LLM + SQL; E = + Kumo Tabular-L fitted on LLM-built features; R = + Kumo Relational (graph layer probe); ER = both. Official test AUROC, mean over repeats (unsubmitted = missing).

| task                    |      D |      E |     ER |      R |
|:------------------------|-------:|-------:|-------:|-------:|
| rel-avito/user-clicks   | 0.5836 | 0.6607 | 0.6268 | 0.6244 |
| rel-avito/user-visits   | 0.6339 | 0.6412 | 0.6616 | 0.6525 |
| rel-event/user-ignore   | 0.5001 | 0.5675 | 0.84   | 0.8767 |
| rel-event/user-repeat   | 0.5559 | 0.7106 | 0.7544 | 0.7939 |
| rel-f1/driver-dnf       | 0.5063 | 0.792  | 0.7437 | 0.7134 |
| rel-f1/driver-top3      | 0.7486 | 0.7838 | 0.863  | 0.863  |
| rel-hm/user-churn       | 0.5951 | 0.5726 | 0.6738 | 0.6738 |
| rel-trial/study-outcome | 0.491  | 0.5914 | 0.6902 | 0.6902 |

Share of episodes that submitted:

| task                    |   D |   E |   ER |   R |
|:------------------------|----:|----:|-----:|----:|
| rel-avito/user-clicks   |   1 | 1   |    1 |   1 |
| rel-avito/user-visits   |   1 | 1   |    1 |   1 |
| rel-event/user-ignore   |   1 | 1   |    1 |   1 |
| rel-event/user-repeat   |   1 | 1   |    1 |   1 |
| rel-f1/driver-dnf       |   1 | 0.5 |    1 |   1 |
| rel-f1/driver-top3      |   1 | 1   |    1 |   1 |
| rel-hm/user-churn       |   1 | 1   |    1 |   1 |
| rel-trial/study-outcome |   1 | 1   |    1 |   1 |

Mean over tasks: {'D': 0.5768, 'E': 0.665, 'ER': 0.7317, 'R': 0.736}
