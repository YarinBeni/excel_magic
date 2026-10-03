# rel-event/user-repeat: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | random | 0.6742 | block_04 | 0.6650 | 0.6265 |  | intrinsic_dim (-0.41) |
| tabpfn | kmeans | 0.6742 | block_03 | 0.6751 | 0.7016 |  | effective_rank (+0.69) |
| tabpfn | label | 0.6742 | block_08 | 0.7330 | 0.7201 | 0.7080 | anisotropy (-0.81) |
| tabpfn-2.5 | random | 0.6742 | block_10 | 0.7072 | 0.5691 |  | anisotropy (+0.31) |
| tabpfn-2.5 | kmeans | 0.6742 | block_09 | 0.7209 | 0.6892 |  | intrinsic_dim (+0.73) |
| tabpfn-2.5 | label | 0.6742 | block_23 | 0.7425 | 0.7425 | 0.6849 | effective_rank (+0.83) |
| tabiclv2 | random | 0.6742 | icl_03 | 0.6895 | 0.6679 |  | effective_rank (-0.85) |
| tabiclv2 | kmeans | 0.6742 | row_emb | 0.7017 | 0.6862 |  | anisotropy (-0.20) |
| tabiclv2 | label | 0.6742 | icl_06 | 0.6843 | 0.6968 | 0.6769 | kmeans_silhouette (+0.84) |
| kumo-s | random | 0.6742 | icl_09 | 0.7130 | 0.7112 |  | anisotropy (-0.21) |
| kumo-s | kmeans | 0.6742 | icl_07 | 0.7138 | 0.7320 |  | kmeans_silhouette (+0.79) |
| kumo-s | label | 0.6742 | final | 0.7045 | 0.7045 | 0.7052 | kmeans_silhouette (+0.79) |
| kumo-m | random | 0.6742 | icl_01 | 0.7181 | 0.6685 |  | effective_rank (-0.82) |
| kumo-m | kmeans | 0.6742 | icl_14 | 0.7070 | 0.7261 |  | effective_rank (-0.85) |
| kumo-m | label | 0.6742 | icl_02 | 0.7236 | 0.7301 | 0.6611 | kmeans_silhouette (+0.56) |
| kumo-l | random | 0.6742 | row_emb | 0.7178 | 0.6888 |  | anisotropy (-0.71) |
| kumo-l | kmeans | 0.6742 | icl_01 | 0.7039 | 0.7113 |  | effective_rank (-0.73) |
| kumo-l | label | 0.6742 | icl_15 | 0.7165 | 0.7107 | 0.6678 | intrinsic_dim (-0.35) |