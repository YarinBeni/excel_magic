# P4: guarded auto-research over the SQL harness (BIRD Arcwise-Plat)

Search 298 / accept 75 / held-out 125 questions; 13 attempts, 1 kept.

Initial {"link": null, "link_k": 20, "triage": false, "triage_threshold": 0.5, "n": 1, "verifier": "greedy"}
Final   {"link": null, "link_k": 20, "triage": true, "triage_threshold": 0.5, "n": 1, "verifier": "greedy"}

Held-out accuracy: initial 0.6800 -> final 0.7200 (paired delta +0.0400, se 0.0176)

| # | knob | value | hypothesis | search delta | se | accept delta | kept |
|---|---|---|---|---|---|---|---|
| 0 | verifier | self_consistency | Using self-consistency verification instead of a greedy appr | -0.0101 | 0.0075 | +0.0267 | no |
| 1 | link | gliclass | Switching to 'gliclass' for link may improve retrieval, pote | -0.0201 | 0.0171 | +0.0000 | no |
| 2 | n | 4 | Increasing the number of generations (n) from 1 to 4 improve | -0.0101 | 0.0058 | +0.0267 | no |
| 3 | verifier | stack | Using the 'stack' verifier will improve performance by maint | -0.0101 | 0.0075 | +0.0267 | no |
| 4 | triage | True | Enabling triage will improve score by filtering out low-conf | +0.0134 | 0.0095 | +0.0133 | yes |
| 5 | triage_threshold | 0.3 | Lowering the triage threshold may improve performance by all | -0.0034 | 0.0058 | +0.0000 | no |
| 6 | link_k | 10 | Decreasing link_k from 20 to 10 may improve search performan | -0.0067 | 0.0067 | -0.0133 | no |
| 7 | link_k | 40 | Increasing link_k from 20 to 40 will improve recall by retri | -0.0067 | 0.0067 | +0.0000 | no |
| 8 | link | rerank | Changing the link mode from null to 'rerank' might improve p | -0.0671 | 0.0204 | -0.0133 | no |
| 9 | n | 8 | Increasing n from 1 to 8 will improve the score by providing | +0.0034 | 0.0058 | -0.0267 | no |
| 10 | n | 16 | Increasing n from 1 to 16 will improve performance by allowi | -0.0034 | 0.0058 | +0.0000 | no |
| 11 | triage_threshold | 0.7 | (fallback: untried value) | +0.0000 | 0.0082 | +0.0000 | no |
| 12 | triage | False | (fallback: untried value) | -0.0134 | 0.0095 | -0.0133 | no |
