| method | val MAP@K x100 |
|---|---|
| GlobalPopularity | 0.367 |
| PastVisit | 1.964 |
| kNN-CF[row] | 0.099 |
| PastVisit+kNN-CF[row] | 1.954 |
| kNN-CF[agg] | 0.217 |
| PastVisit+kNN-CF[agg] | 1.954 |
| kNN-CF[tabpfn_agg_kmeans] | 0.136 |
| PastVisit+kNN-CF[tabpfn_agg_kmeans] | 1.954 |

queries: val=3000; hist_days=365; k_neighbors=50; published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93