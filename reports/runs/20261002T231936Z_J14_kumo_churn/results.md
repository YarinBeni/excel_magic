| method | AUROC |
|---|---|
| MajorityPrior | 0.5000 |
| Supervised[hgb on agg] | 0.6725 |
| Supervised[tabpfn on agg] | 0.6715 |
| kNN[agg] | 0.6530 |
| kNN[tabpfn_agg_kmeans] | 0.6484 |
| kNN[kumo_relational_kmeans] | 0.6443 |
| kNN[kumo_relational_random] | 0.6603 |

customers=20000 at 2020-08-31, pos_rate=0.817, hist_days=365, k=50, 5-fold over customers (not the official temporal test split)