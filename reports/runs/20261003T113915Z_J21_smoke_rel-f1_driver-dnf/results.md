# rel-f1/driver-dnf: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | label | 0.7703 | block_11 | 0.6145 | 0.6145 | 0.8189 | kmeans_silhouette (-0.55) |
| kumo-s | label | 0.7703 | row_emb | 0.7546 | 0.7391 | 0.8176 | effective_rank (-0.68) |