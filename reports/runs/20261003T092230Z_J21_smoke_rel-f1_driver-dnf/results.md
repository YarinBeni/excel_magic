# rel-f1/driver-dnf: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | label | 0.8200 | block_11 | 0.7444 | 0.7444 | 0.7779 | anisotropy (-0.45) |
| kumo-s | label | 0.8200 | row_emb | 0.7504 | 0.6354 | 0.8080 | intrinsic_dim (-0.74) |