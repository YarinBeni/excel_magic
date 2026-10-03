# rel-f1/driver-dnf: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.7901 | block_00 | 0.8018 | 0.7437 |  | intrinsic_dim (-0.57) |
| tabpfn | kmeans | 0.7901 | block_10 | 0.7440 | 0.7143 |  | effective_rank (-0.27) |
| tabpfn | label | 0.7901 | block_08 | 0.8237 | 0.7748 | 0.7959 | anisotropy (-0.61) |
| tabpfn-2.5 | random | 0.7901 | block_04 | 0.8098 | 0.7786 |  | effective_rank (-0.45) |
| tabpfn-2.5 | kmeans | 0.7901 | block_06 | 0.7801 | 0.7701 |  | anisotropy (+0.43) |
| tabpfn-2.5 | label | 0.7901 | block_04 | 0.8096 | 0.7982 | 0.8246 | anisotropy (-0.24) |
| tabiclv2 | random | 0.7901 | row_emb | 0.7756 | 0.7146 |  | kmeans_silhouette (+0.16) |
| tabiclv2 | kmeans | 0.7901 | row_emb | 0.7643 | 0.7572 |  | anisotropy (-0.80) |
| tabiclv2 | label | 0.7901 | row_emb | 0.7920 | 0.7029 | 0.8098 | anisotropy (-0.48) |
| kumo-s | random | 0.7901 | row_emb | 0.5637 | 0.6990 |  | intrinsic_dim (+0.68) |
| kumo-s | kmeans | 0.7901 | row_emb | 0.7698 | 0.7448 |  | effective_rank (-0.81) |
| kumo-s | label | 0.7901 | row_emb | 0.5759 | 0.6873 | 0.8162 | anisotropy (+0.72) |
| kumo-m | random | 0.7901 | icl_17 | 0.7223 | 0.7132 |  | anisotropy (+0.71) |
| kumo-m | kmeans | 0.7901 | final | 0.7820 | 0.7820 |  | effective_rank (+0.29) |
| kumo-m | label | 0.7901 | icl_20 | 0.7694 | 0.7795 | 0.8051 | effective_rank (-0.50) |
| kumo-l | random | 0.7901 | icl_00 | 0.7154 | 0.6106 |  | effective_rank (-0.55) |
| kumo-l | kmeans | 0.7901 | final | 0.7588 | 0.7588 |  | effective_rank (-0.80) |
| kumo-l | label | 0.7901 | icl_00 | 0.7061 | 0.6326 | 0.8085 | effective_rank (-0.49) |