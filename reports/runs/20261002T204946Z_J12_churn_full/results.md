| method | AUROC |
|---|---|
| MajorityPrior | 0.5000 |
| Supervised[hgb on agg] | 0.6725 |
| Supervised[tabpfn on agg] | 0.6715 |
| kNN[row] | 0.5198 |
| kNN[agg] | 0.6530 |
| kNN[svd] | 0.5901 |
| kNN[svd_agg] | 0.6375 |
| kNN[tabpfn_agg_kmeans] | 0.6484 |
| kNN[tabpfn_agg_random] | 0.6425 |
| kNN[tabpfn_svd_kmeans] | 0.6450 |

customers=20000 at 2020-08-31, pos_rate=0.817, hist_days=365, k=50, 5-fold over customers (not the official temporal test split)