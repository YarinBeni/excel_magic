# P0 verifier study: 497 BIRD Arcwise-Plat questions, 4473 SQL candidates

Greedy accuracy 0.692; oracle (any candidate correct) 0.746; 0.690 of candidates correct.

| signal | AUROC | within-question AUROC | best-of-N accuracy | ECE raw | ECE isotonic |
|---|---|---|---|---|---|
| query executes | 0.556 | 0.631 | 0.706 | 0.276 | 0.000 |
| non-empty result | 0.600 | 0.657 | 0.706 | 0.248 | 0.001 |
| generator logprob | 0.579 | 0.512 | 0.692 | nan | 0.005 |
| self-consistency | 0.651 | 0.721 | 0.706 | 0.219 | 0.009 |
| GLiClass large v3 (zero-shot) | 0.685 | 0.695 | 0.712 | 0.073 | 0.018 |
| NLI cross-encoder (DeBERTa-v3-large) | 0.656 | 0.733 | 0.720 | 0.328 | 0.050 |
| LLM judge (2nd family) | 0.786 | 0.748 | 0.714 | 0.197 | 0.047 |
| stack: no-model signals | 0.676 | 0.725 | 0.706 | 0.026 | 0.092 |
| stack: + GLiClass (+ taxonomy) | 0.725 | 0.766 | 0.712 | 0.033 | 0.064 |
| stack: + GLiClass + NLI | 0.722 | 0.765 | 0.710 | 0.036 | 0.071 |
| stack: + judge | 0.791 | 0.762 | 0.714 | 0.023 | 0.051 |

Cascade (stack with GLiClass; the most uncertain share goes to the judge):

| judge share | AUROC | best-of-N |
|---|---|---|
| 0% | 0.725 | 0.712 |
| 10% | 0.725 | 0.716 |
| 20% | 0.729 | 0.712 |
| 30% | 0.735 | 0.708 |
| 50% | 0.747 | 0.704 |
| 100% | 0.786 | 0.714 |

Schema linking (recall of the gold SQL's columns among the top-k ranked columns):

| scorer            |   recall@5 |   recall@10 |   recall@20 |   all@10 |   all@20 |
|:------------------|-----------:|------------:|------------:|---------:|---------:|
| embed_bge_small   |      0.503 |       0.645 |       0.774 |    0.275 |    0.472 |
| gliclass_large_v3 |      0.648 |       0.789 |       0.877 |    0.424 |    0.618 |
| gliner_bi_base_v2 |      0.24  |       0.35  |       0.53  |    0.082 |    0.213 |
| lexical           |      0.576 |       0.706 |       0.811 |    0.317 |    0.468 |
| rerank_bge_m3     |      0.541 |       0.687 |       0.802 |    0.325 |    0.52  |

Latency (seconds per item): gliclass_s_per_item 0.0054, gliclass_taxonomy_s_per_item 0.0059, nli_s_per_item 0.0052, judge_s_per_item_wall 0.0059, judge_fallback_share 0.0000, judge_s_per_call_sequential 0.0639, lexical 0.0003, embed_bge_small 0.0150, rerank_bge_m3 0.1049, gliclass_large_v3 0.0496, gliner_bi_base_v2 0.0485
