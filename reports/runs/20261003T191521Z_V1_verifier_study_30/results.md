# P0 verifier study: 30 BIRD Arcwise-Plat questions, 270 SQL candidates

Greedy accuracy 0.667; oracle (any candidate correct) 0.767; 0.667 of candidates correct.

| signal | AUROC | within-question AUROC | best-of-N accuracy | ECE raw | ECE isotonic |
|---|---|---|---|---|---|
| query executes | 0.500 | 0.500 | 0.700 | 0.333 | 0.074 |
| non-empty result | 0.506 | 0.521 | 0.700 | 0.330 | 0.072 |
| generator logprob | 0.564 | 0.679 | 0.700 | nan | 0.143 |
| self-consistency | 0.723 | 0.516 | 0.700 | 0.261 | 0.210 |
| GLiClass large v3 (zero-shot) | 0.569 | 0.271 | 0.633 | 0.155 | 0.088 |
| NLI cross-encoder (DeBERTa-v3-large) | 0.356 | 0.450 | 0.700 | 0.379 | 0.099 |
| LLM judge (2nd family) | 0.645 | 0.544 | 0.700 | 0.249 | 0.125 |
| stack: no-model signals | 0.582 | 0.519 | 0.700 | 0.179 | 0.238 |
| stack: + GLiClass (+ taxonomy) | 0.554 | 0.123 | 0.633 | 0.269 | 0.255 |
| stack: + GLiClass + NLI | 0.563 | 0.095 | 0.633 | 0.248 | 0.255 |
| stack: + judge | 0.527 | 0.070 | 0.633 | 0.288 | 0.263 |

Cascade (stack with GLiClass; the most uncertain share goes to the judge):

| judge share | AUROC | best-of-N |
|---|---|---|
| 0% | 0.554 | 0.633 |
| 10% | 0.539 | 0.633 |
| 20% | 0.537 | 0.633 |
| 30% | 0.529 | 0.633 |
| 50% | 0.551 | 0.633 |
| 100% | 0.645 | 0.700 |

Schema linking (recall of the gold SQL's columns among the top-k ranked columns):

| scorer            |   recall@5 |   recall@10 |   recall@20 |   all@10 |   all@20 |
|:------------------|-----------:|------------:|------------:|---------:|---------:|
| embed_bge_small   |      0.697 |       0.874 |       1     |    0.533 |    1     |
| gliclass_large_v3 |      0.705 |       0.859 |       0.989 |    0.533 |    0.967 |
| gliner_bi_base_v2 |      0.416 |       0.609 |       0.924 |    0.167 |    0.667 |
| lexical           |      0.68  |       0.787 |       0.983 |    0.533 |    0.9   |
| rerank_bge_m3     |      0.647 |       0.863 |       1     |    0.533 |    1     |

Latency (seconds per item): gliclass_s_per_item 0.0088, gliclass_taxonomy_s_per_item 0.0066, nli_s_per_item 0.0080, judge_s_per_item_wall 0.0064, judge_fallback_share 0.0000, judge_s_per_call_sequential 0.0547, lexical 0.0001, embed_bge_small 0.0192, rerank_bge_m3 0.0325, gliclass_large_v3 0.0379, gliner_bi_base_v2 0.2983
