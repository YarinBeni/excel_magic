# rel-f1/driver-dnf: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.7951 | block_02 | 0.7845 | 0.6430 |  | anisotropy (+0.57) |
| tabpfn | kmeans | 0.7951 | block_07 | 0.7895 | 0.7969 |  | anisotropy (-0.33) |
| tabpfn | label | 0.7951 | block_09 | 0.8081 | 0.7471 | 0.8167 | intrinsic_dim (+0.46) |
| tabpfn-2.5 | random | 0.7951 | block_09 | 0.7922 | 0.7395 |  | effective_rank (-0.71) |
| tabpfn-2.5 | kmeans | 0.7951 | block_04 | 0.7475 | 0.7694 |  | kmeans_silhouette (-0.34) |
| tabpfn-2.5 | label | 0.7951 | block_14 | 0.8070 | 0.7236 | 0.8093 | anisotropy (+0.23) |
| tabiclv2 | random | 0.7951 | row_emb | 0.8085 | 0.7639 |  | intrinsic_dim (+0.39) |
| tabiclv2 | kmeans | 0.7951 | row_emb | 0.8142 | 0.7491 |  | effective_rank (-0.85) |
| tabiclv2 | label | 0.7951 | icl_00 | 0.8309 | 0.4002 | 0.8105 | effective_rank (-0.89) |
| kumo-s | random | 0.7951 | icl_01 | 0.7990 | 0.7033 |  | effective_rank (-0.89) |
| kumo-s | kmeans | 0.7951 | final | 0.8006 | 0.8006 |  | effective_rank (-0.72) |
| kumo-s | label | 0.7951 | row_emb | 0.8100 | 0.4303 | 0.8102 | effective_rank (-0.83) |
| kumo-m | random | 0.7951 | icl_00 | 0.7804 | 0.7064 |  | effective_rank (-0.70) |
| kumo-m | kmeans | 0.7951 | row_emb | 0.7883 | 0.7753 |  | intrinsic_dim (-0.61) |
| kumo-m | label | 0.7951 | icl_00 | 0.7037 | 0.5054 | 0.7993 | anisotropy (-0.34) |
| kumo-l | random | 0.7951 | icl_14 | 0.7523 | 0.6360 |  | intrinsic_dim (+0.61) |
| kumo-l | kmeans | 0.7951 | final | 0.7528 | 0.7528 |  | anisotropy (+0.55) |
| kumo-l | label | 0.7951 | icl_00 | 0.7350 | 0.6734 | 0.8147 | effective_rank (-0.68) |