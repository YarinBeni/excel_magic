# rel-f1/driver-top3: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.8912 | block_00 | 0.8987 | 0.7644 |  | anisotropy (+0.34) |
| tabpfn | kmeans | 0.8912 | block_08 | 0.8802 | 0.8738 |  | anisotropy (-0.66) |
| tabpfn | label | 0.8912 | block_00 | 0.8990 | 0.9039 | 0.8849 | intrinsic_dim (-0.76) |
| tabpfn-2.5 | random | 0.8912 | block_00 | 0.8809 | 0.8858 |  | anisotropy (-0.21) |
| tabpfn-2.5 | kmeans | 0.8912 | block_18 | 0.8995 | 0.8953 |  | intrinsic_dim (-0.67) |
| tabpfn-2.5 | label | 0.8912 | block_00 | 0.8810 | 0.8889 | 0.8919 | anisotropy (-0.57) |
| tabiclv2 | random | 0.8912 | icl_09 | 0.7739 | 0.8261 |  | intrinsic_dim (-0.79) |
| tabiclv2 | kmeans | 0.8912 | row_emb | 0.9019 | 0.8810 |  | effective_rank (-0.90) |
| tabiclv2 | label | 0.8912 | icl_09 | 0.8988 | 0.9023 | 0.8666 | kmeans_silhouette (-0.78) |
| kumo-s | random | 0.8912 | icl_01 | 0.8864 | 0.8126 |  | effective_rank (-0.90) |
| kumo-s | kmeans | 0.8912 | final | 0.8900 | 0.8900 |  | intrinsic_dim (-0.90) |
| kumo-s | label | 0.8912 | icl_02 | 0.8974 | 0.8841 | 0.8408 | effective_rank (-0.80) |
| kumo-m | random | 0.8912 | icl_11 | 0.8616 | 0.8295 |  | effective_rank (+0.38) |
| kumo-m | kmeans | 0.8912 | icl_00 | 0.9123 | 0.8448 |  | kmeans_silhouette (-0.81) |
| kumo-m | label | 0.8912 | icl_04 | 0.8672 | 0.8256 | 0.8652 | effective_rank (-0.39) |
| kumo-l | random | 0.8912 | row_emb | 0.8861 | 0.8300 |  | kmeans_silhouette (-0.72) |
| kumo-l | kmeans | 0.8912 | icl_21 | 0.8944 | 0.8902 |  | intrinsic_dim (-0.80) |
| kumo-l | label | 0.8912 | icl_11 | 0.9024 | 0.8940 | 0.8610 | anisotropy (+0.49) |