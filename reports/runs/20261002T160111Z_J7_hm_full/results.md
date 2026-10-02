| method | val MAP@K x100 | test MAP@K x100 |
|---|---|---|
| GlobalPopularity | 0.342 | 0.292 |
| PastVisit | 1.904 | 2.199 |
| kNN-CF[row] | 0.136 | 0.144 |
| PastVisit+kNN-CF[row] | 1.897 | 2.191 |
| kNN-CF[agg] | 0.258 | 0.259 |
| PastVisit+kNN-CF[agg] | 1.897 | 2.191 |
| kNN-CF[tabpfn_agg_kmeans] | 0.248 | 0.227 |
| PastVisit+kNN-CF[tabpfn_agg_kmeans] | 1.897 | 2.191 |
| kNN-CF[tabpfn_agg_random] | 0.207 | 0.170 |
| PastVisit+kNN-CF[tabpfn_agg_random] | 1.897 | 2.191 |
| kNN-CF[tabpfn_row_kmeans] | 0.135 | 0.140 |
| PastVisit+kNN-CF[tabpfn_row_kmeans] | 1.897 | 2.191 |

queries: val=74575, test=67144; hist_days=365; k_neighbors=50; published test rows (x100): GlobalPop 0.30, PastVisit 0.89, LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93