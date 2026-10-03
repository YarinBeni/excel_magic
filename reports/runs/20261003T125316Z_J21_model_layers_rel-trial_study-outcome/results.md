# rel-trial/study-outcome: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.6302 | block_01 | 0.6209 | 0.6282 |  | kmeans_silhouette (-0.86) |
| tabpfn | kmeans | 0.6302 | block_09 | 0.5807 | 0.5725 |  | effective_rank (-0.30) |
| tabpfn | label | 0.6302 | block_05 | 0.6674 | 0.7022 | 0.6890 | effective_rank (+0.71) |
| tabpfn-2.5 | random | 0.6302 | block_03 | 0.6345 | 0.6485 |  | intrinsic_dim (+0.66) |
| tabpfn-2.5 | kmeans | 0.6302 | block_20 | 0.6192 | 0.6074 |  | kmeans_silhouette (-0.54) |
| tabpfn-2.5 | label | 0.6302 | block_22 | 0.7120 | 0.7116 | 0.6782 | anisotropy (-0.52) |
| tabiclv2 | random | 0.6302 | row_emb | 0.7099 | 0.6774 |  | effective_rank (-0.80) |
| tabiclv2 | kmeans | 0.6302 | icl_00 | 0.7015 | 0.6837 |  | effective_rank (-0.75) |
| tabiclv2 | label | 0.6302 | icl_00 | 0.6984 | 0.6887 | 0.6849 | kmeans_silhouette (+0.71) |
| kumo-s | random | 0.6302 | row_emb | 0.7000 | 0.6719 |  | effective_rank (-0.84) |
| kumo-s | kmeans | 0.6302 | icl_00 | 0.6744 | 0.6594 |  | effective_rank (-0.93) |
| kumo-s | label | 0.6302 | row_emb | 0.7071 | 0.6811 | 0.6826 | effective_rank (-0.80) |
| kumo-m | random | 0.6302 | icl_00 | 0.6986 | 0.6689 |  | anisotropy (-0.79) |
| kumo-m | kmeans | 0.6302 | row_emb | 0.6754 | 0.6535 |  | kmeans_silhouette (-0.76) |
| kumo-m | label | 0.6302 | icl_01 | 0.7117 | 0.6839 | 0.6850 | effective_rank (-0.68) |
| kumo-l | random | 0.6302 | row_emb | 0.6870 | 0.6684 |  | effective_rank (-0.83) |
| kumo-l | kmeans | 0.6302 | row_emb | 0.6700 | 0.6524 |  | effective_rank (-0.71) |
| kumo-l | label | 0.6302 | row_emb | 0.6880 | 0.6704 | 0.6882 | effective_rank (-0.93) |