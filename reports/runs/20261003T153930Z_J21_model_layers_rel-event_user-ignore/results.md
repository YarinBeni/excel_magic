# rel-event/user-ignore: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.8834 | block_11 | 0.8061 | 0.8061 |  | intrinsic_dim (-0.76) |
| tabpfn | kmeans | 0.8834 | block_09 | 0.7872 | 0.7776 |  | anisotropy (-0.83) |
| tabpfn | label | 0.8834 | block_09 | 0.8934 | 0.9008 | 0.8916 | effective_rank (+0.88) |
| tabpfn-2.5 | random | 0.8834 | block_08 | 0.8292 | 0.7878 |  | anisotropy (+0.69) |
| tabpfn-2.5 | kmeans | 0.8834 | block_11 | 0.8078 | 0.8171 |  | effective_rank (+0.27) |
| tabpfn-2.5 | label | 0.8834 | block_14 | 0.9008 | 0.9058 | 0.8902 | anisotropy (-0.46) |
| tabiclv2 | random | 0.8834 | icl_05 | 0.8635 | 0.8651 |  | intrinsic_dim (-0.79) |
| tabiclv2 | kmeans | 0.8834 | icl_06 | 0.8735 | 0.8731 |  | intrinsic_dim (-0.81) |
| tabiclv2 | label | 0.8834 | icl_05 | 0.8677 | 0.8478 | 0.8709 | effective_rank (-0.93) |
| kumo-s | random | 0.8834 | icl_01 | 0.8908 | 0.8667 |  | effective_rank (-0.94) |
| kumo-s | kmeans | 0.8834 | icl_02 | 0.8963 | 0.8788 |  | anisotropy (+0.94) |
| kumo-s | label | 0.8834 | row_emb | 0.8876 | 0.8703 | 0.8886 | effective_rank (-0.89) |
| kumo-m | random | 0.8834 | icl_00 | 0.8797 | 0.8311 |  | effective_rank (-0.79) |
| kumo-m | kmeans | 0.8834 | icl_03 | 0.8806 | 0.8868 |  | kmeans_silhouette (+0.72) |
| kumo-m | label | 0.8834 | icl_04 | 0.8789 | 0.8431 | 0.8820 | intrinsic_dim (-0.85) |
| kumo-l | random | 0.8834 | row_emb | 0.8934 | 0.8263 |  | effective_rank (-0.93) |
| kumo-l | kmeans | 0.8834 | icl_02 | 0.8960 | 0.8921 |  | effective_rank (-0.80) |
| kumo-l | label | 0.8834 | icl_20 | 0.8296 | 0.8313 | 0.8920 | effective_rank (-0.78) |