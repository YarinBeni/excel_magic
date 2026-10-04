# Q3: LLM hypotheses on top of the frozen relational FM (RelBench, official test AUROC)

| model                        | task                    |   system |     fm |    gain |   accepted |   tested |
|:-----------------------------|:------------------------|---------:|-------:|--------:|-----------:|---------:|
| Qwen3-Coder-30B-A3B-Instruct | rel-avito/user-clicks   |   0.6036 | 0.6235 | -0.02   |        1.5 |     10   |
| Qwen3-Coder-30B-A3B-Instruct | rel-avito/user-visits   |   0.6516 | 0.6516 |  0      |        0   |      9   |
| Qwen3-Coder-30B-A3B-Instruct | rel-event/user-ignore   |   0.8767 | 0.8767 |  0      |        0   |      9.5 |
| Qwen3-Coder-30B-A3B-Instruct | rel-event/user-repeat   |   0.7909 | 0.7909 |  0      |        0   |      9.5 |
| Qwen3-Coder-30B-A3B-Instruct | rel-f1/driver-dnf       |   0.7518 | 0.7486 |  0.0032 |        0.5 |      5   |
| Qwen3-Coder-30B-A3B-Instruct | rel-f1/driver-top3      |   0.9272 | 0.863  |  0.0642 |        1.5 |     10   |
| Qwen3-Coder-30B-A3B-Instruct | rel-hm/user-churn       |   0.6738 | 0.6738 |  0      |        0   |      9.5 |
| Qwen3-Coder-30B-A3B-Instruct | rel-trial/study-outcome |   0.6919 | 0.6902 |  0.0016 |        0.5 |     10   |

Qwen3-Coder-30B-A3B-Instruct: mean system 0.7459 vs FM 0.7398; gain +0.0061; better on 3 / worse on 1 of 8 tasks
