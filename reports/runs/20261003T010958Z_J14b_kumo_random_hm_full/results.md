| method | val MAP@K x100 | test MAP@K x100 |
|---|---|---|
| GlobalPopularity | 0.342 | 0.292 |
| PastVisit | 1.904 | 2.199 |
| kNN-CF[purchase_matrix] | 1.164 | 1.203 |
| Past+kNN-CF[purchase_matrix] | 1.919 | 2.216 |
| ItemKNN | 1.034 | 1.073 |
| Past+ItemKNN | 1.901 | 2.197 |
| kNN-CF[agg] | 0.258 | 0.259 |
| Past+kNN-CF[agg] | 1.889 | 2.184 |
| kNN-CF[kumo_relational_random] | 0.253 | 0.244 |
| Past+kNN-CF[kumo_relational_random] | 1.886 | 2.182 |

queries: val=74575, test=67144; hist_days=365; k_neighbors=50; published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93