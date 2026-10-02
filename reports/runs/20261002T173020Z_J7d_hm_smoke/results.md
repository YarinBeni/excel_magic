| method | val MAP@K x100 |
|---|---|
| GlobalPopularity | 0.367 |
| PastVisit | 1.964 |
| kNN-CF[purchase_matrix] | 1.004 |
| Past+kNN-CF[purchase_matrix] | 1.970 |
| ItemKNN | 0.879 |
| Past+ItemKNN | 1.928 |
| kNN-CF[svd] | 0.535 |
| Past+kNN-CF[svd] | 1.920 |
| kNN-CF[tabpfn_svd_kmeans] | 0.179 |
| Past+kNN-CF[tabpfn_svd_kmeans] | 1.924 |

queries: val=3000; hist_days=365; k_neighbors=50; published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93