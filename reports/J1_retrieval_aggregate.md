ERR runs/20261002T154754Z_J1_retrieval_seed0 tabpfn_agg_kmeans TypeError: Got unsupported ScalarType BFloat16
ERR runs/20261002T154754Z_J1_retrieval_seed0 tabpfn_agg TypeError: Got unsupported ScalarType BFloat16
ERR runs/20261002T154911Z_J1_retrieval_seed1 tabpfn_agg_kmeans TypeError: Got unsupported ScalarType BFloat16
ERR runs/20261002T154911Z_J1_retrieval_seed1 tabpfn_agg TypeError: Got unsupported ScalarType BFloat16
ERR runs/20261002T155020Z_J1_retrieval_seed2 tabpfn_agg_kmeans TypeError: Got unsupported ScalarType BFloat16
ERR runs/20261002T155020Z_J1_retrieval_seed2 tabpfn_agg TypeError: Got unsupported ScalarType BFloat16

### synth_shop (3 seed(s), 1151 query entities; popularity baseline MAP@10 = 0.034, novel-only popularity = 0.040)

| embedder | seg P@10 (mean +/- sd) | seg 10-NN acc | fut MAP@10 (mean +/- sd) | fut recall@10 | novel MAP@10 (mean +/- sd) | s |
|---|---|---|---|---|---|---|
| row | 0.161 +/- 0.000 | 0.146 | 0.022 +/- 0.000 | 0.061 | 0.024 +/- 0.000 | 0 |
| agg | 0.726 +/- 0.000 | 0.842 | 0.048 +/- 0.000 | 0.142 | 0.054 +/- 0.000 | 0 |
| tabpfn_row | 0.169 +/- 0.005 | 0.171 | 0.021 +/- 0.001 | 0.060 | nan +/- 0.000 | 6 |
| tabpfn_agg | 0.492 +/- 0.344 | 0.567 | 0.036 +/- 0.013 | 0.108 | nan +/- 0.000 | 18 |
| tabpfn_agg_rand4 | 0.525 +/- 0.325 | 0.617 | 0.037 +/- 0.012 | 0.107 | 0.044 +/- 0.014 | 80 |
| tabpfn_agg_kmeans | 0.824 +/- 0.021 | 0.873 | 0.048 +/- 0.001 | 0.139 | 0.058 +/- 0.001 | 20 |
| tabpfn_agg_feat4 | 0.680 +/- 0.030 | 0.787 | 0.042 +/- 0.001 | 0.121 | 0.048 +/- 0.001 | 72 |
| tabpfn_agg_churn | 0.182 +/- 0.000 | 0.173 | 0.022 +/- 0.000 | 0.061 | nan +/- 0.000 | 15 |
| gnn | 0.877 +/- 0.007 | 0.902 | 0.047 +/- 0.001 | 0.140 | 0.056 +/- 0.003 | 1 |
| openrfm | 0.249 +/- 0.000 | 0.306 | 0.032 +/- 0.000 | 0.095 | 0.039 +/- 0.000 | 8 |
| openrfm_ctx16 | 0.237 +/- 0.000 | 0.291 | 0.034 +/- 0.000 | 0.089 | nan +/- 0.000 | 428 |
| openrfm_random | 0.168 +/- 0.000 | 0.169 | 0.013 +/- 0.000 | 0.042 | nan +/- 0.000 | 16 |
| kumo_relational | 0.330 +/- 0.132 | 0.422 | 0.036 +/- 0.009 | 0.096 | 0.041 +/- 0.009 | 46 |
