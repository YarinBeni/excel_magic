| method | val MAP@K x100 | test MAP@K x100 |
|---|---|---|
| GlobalPopularity | 0.342 | 0.292 |
| PastVisit | 1.904 | 2.199 |
| kNN-CF[purchase_matrix] | 1.162 | 1.200 |
| Past+kNN-CF[purchase_matrix] | 1.919 | 2.216 |
| ItemKNN | 1.035 | 1.073 |
| Past+ItemKNN | 1.901 | 2.197 |
| kNN-CF[svd] | 0.822 | 0.833 |
| Past+kNN-CF[svd] | 1.894 | 2.195 |
| kNN-CF[svd_agg] | 0.661 | 0.654 |
| Past+kNN-CF[svd_agg] | 1.888 | 2.185 |
| kNN-CF[tabpfn_svd_kmeans] | 0.401 | 0.419 |
| Past+kNN-CF[tabpfn_svd_kmeans] | 1.886 | 2.180 |
| kNN-CF[tabpfn_svd_random] | 0.514 | 0.475 |
| Past+kNN-CF[tabpfn_svd_random] | 1.885 | 2.182 |
| kNN-CF[agg] | 0.258 | 0.259 |
| Past+kNN-CF[agg] | 1.889 | 2.184 |
| kNN-CF[tabpfn_agg_kmeans] | 0.248 | 0.227 |
| Past+kNN-CF[tabpfn_agg_kmeans] | 1.889 | 2.187 |

queries: val=74575, test=67144; hist_days=365; k_neighbors=50; published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93