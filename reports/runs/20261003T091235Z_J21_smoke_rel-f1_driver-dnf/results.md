# rel-f1/driver-dnf: every model, every layer (test AUROC, official evaluator)

| model | target | raw features | best layer by val | its test | last layer test | own prediction | geometry best predictor (rho) |
|---|---|---|---|---|---|---|---|
| tabpfn | label | 0.8200 | block_11 | 0.5567 | 0.5567 | 0.7913 | anisotropy (-0.51) |
| kumo-s | label | 0.8200 | row_emb | 0.6834 | 0.6118 | 0.8069 | effective_rank (-0.59) |