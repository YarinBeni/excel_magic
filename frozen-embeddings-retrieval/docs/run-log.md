# Run log (append-only). Every number comes from runs/<id>/metrics.json.

## 2026-10-02 seed-0 results (migrated from the tabfm-lab prototype)

Synthetic shop DB (2000 customers, 300 products, 6 hidden segments).

| embedder | dim | seg P@10 | seg MAP@10 | seg kNN acc | fut MAP@10 | fut recall@10 | fut hit@10 | s |
|---|---|---|---|---|---|---|---|---|
| row | 10 | 0.161 | 0.064 | 0.146 | 0.022 | 0.061 | 0.215 | 0.5 |
| agg | 23 | 0.726 | 0.646 | 0.843 | 0.048 | 0.142 | 0.447 | 0.3 |
| tabpfn_row | 192 | 0.175 | 0.071 | 0.184 | 0.022 | 0.062 | 0.219 | 6.1 |
| tabpfn_agg | 192 | 0.875 | 0.845 | 0.901 | 0.049 | 0.146 | 0.454 | 14.7 |
| tabpfn_agg_churn | 192 | 0.182 | 0.081 | 0.173 | 0.022 | 0.061 | 0.212 | 14.8 |
| gnn | 64 | 0.885 | 0.855 | 0.907 | 0.049 | 0.140 | 0.442 | 1.7 |
| openrfm | 256 | 0.249 | 0.131 | 0.306 | 0.032 | 0.095 | 0.295 | 15.1 |
| openrfm_ctx16 | 256 | 0.237 | 0.119 | 0.291 | 0.034 | 0.089 | 0.274 | 427.9 |
| openrfm_random | 256 | 0.168 | 0.067 | 0.169 | 0.013 | 0.042 | 0.171 | 14.4 |
| popularity baseline | - | 0.167 (chance) | | | 0.034 | 0.089 | | |

runs: `20261002T063222Z_exp04_relational_retrieval`, `20261002T065947Z_exp04_with_tabpfn_embeddings`.

## Aggregated (runs: 20261002T072442Z_bench_synth_seed0, 20261002T072613Z_bench_synth_seed1, 20261002T072724Z_bench_synth_seed2, 20261002T073608Z_bench_synth_targets_seed0, 20261002T073845Z_bench_synth_targets_seed1, 20261002T074119Z_bench_synth_targets_seed2)

### synth_shop (3 seed(s), 1151 query entities; popularity baseline MAP@10 = 0.034, novel-only popularity = 0.040)

| embedder | seg P@10 (mean +/- sd) | seg 10-NN acc | fut MAP@10 (mean +/- sd) | fut recall@10 | novel MAP@10 (mean +/- sd) | s |
|---|---|---|---|---|---|---|
| row | 0.161 +/- 0.000 | 0.146 | 0.022 +/- 0.000 | 0.061 | 0.024 +/- 0.000 | 1 |
| agg | 0.726 +/- 0.000 | 0.842 | 0.048 +/- 0.000 | 0.142 | 0.054 +/- 0.000 | 0 |
| tabpfn_row | 0.169 +/- 0.005 | 0.171 | 0.021 +/- 0.001 | 0.060 | 0.023 +/- 0.001 | 6 |
| tabpfn_agg | 0.492 +/- 0.344 | 0.567 | 0.036 +/- 0.013 | 0.108 | 0.043 +/- 0.017 | 18 |
| tabpfn_agg_rand4 | 0.525 +/- 0.325 | 0.617 | 0.037 +/- 0.012 | 0.107 | 0.044 +/- 0.014 | 80 |
| tabpfn_agg_kmeans | 0.824 +/- 0.021 | 0.873 | 0.048 +/- 0.001 | 0.139 | 0.058 +/- 0.001 | 20 |
| tabpfn_agg_feat4 | 0.680 +/- 0.030 | 0.787 | 0.042 +/- 0.001 | 0.121 | 0.048 +/- 0.001 | 72 |
| gnn | 0.877 +/- 0.007 | 0.902 | 0.047 +/- 0.001 | 0.141 | 0.056 +/- 0.003 | 2 |
| openrfm | 0.249 +/- 0.000 | 0.306 | 0.032 +/- 0.000 | 0.095 | 0.039 +/- 0.000 | 23 |
| openrfm_random | 0.168 +/- 0.000 | 0.169 | 0.013 +/- 0.000 | 0.042 | 0.014 +/- 0.000 | 16 |

## 2026-10-02 cluster J1b (H200, seeds 0-2): first real Kumo Relational numbers
Synthetic shop DB, segment retrieval P@10 / 10-NN acc (chance 0.167); per seed 0 / 1 / 2.

| embedder | seg P@10 | 10-NN acc | fut MAP@10 |
|---|---|---|---|
| kumo_relational (random in-context target, 512-d) | 0.482 / 0.260 / 0.247 | 0.60 / 0.34 / 0.32 | 0.046 / 0.032 / 0.030 |
| **kumo_relational_kmeans** (k-means pseudo-labels as target) | **0.667 / 0.564 / 0.615** | 0.73 / 0.66 / 0.70 | 0.049 / 0.041 / 0.045 |
| tabpfn_agg_kmeans (GPU, float32) | 0.827 / 0.803 / 0.841 | 0.87 / 0.86 / 0.89 | 0.047 / 0.050 / 0.047 |
| tabpfn_agg (random target) | 0.875 / 0.214 / 0.389 | | |
| gnn (trained) | 0.885 / 0.872 / 0.876 | | |
| openrfm | 0.249 | | |
| agg | 0.726 | 0.843 | 0.048 |

Reading: the real relational FM (NVIDIA KumoRelational, a KumoRFM-2 adaptation) gives usable graph embeddings only when
the in-context target is structure-preserving: k-means target 0.62 +/- 0.05 vs random target 0.33 +/- 0.13. It sits
below the frozen TabPFN over flattened aggregates (0.82) and the GNN (0.88) and below the raw aggregates (0.73) on this
synthetic DB, far above the OpenRFM reproduction (0.25). Source: cluster branch `reports/logs/J1_smoke_50265.log`,
runs `J1_retrieval_seed{0,1,2}`.
