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

## 2026-10-02 cluster J7b: RelBench rel-hm / user-item-purchase, official evaluator (MAP@12 x100)
hist_days=365, k_neighbors=50, val 74,575 queries, test 67,144 queries. Runs: `J7_hm_full` (cluster branch `reports/runs`).

| method | val | test | published test |
|---|---|---|---|
| GlobalPopularity | 0.342 | 0.292 | 0.30 |
| PastVisit (365-day history, most recent first) | 1.904 | 2.199 | 0.89 (paper's variant) |
| kNN-CF[row] | 0.136 | 0.144 | |
| kNN-CF[agg] | 0.258 | 0.259 | |
| kNN-CF[tabpfn_agg_kmeans] | 0.248 | 0.227 | |
| kNN-CF[tabpfn_agg_random] | 0.207 | 0.170 | |
| kNN-CF[tabpfn_row_kmeans] | 0.135 | 0.140 | |
| PastVisit + kNN fill (any) | 1.897 | 2.191 | |
| published: LightGBM 0.38, GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93 | | | |

Reading: NEGATIVE for H1/H3 on real data. Customer embeddings built from coarse aggregates (product-group mix, recency,
spend) do not encode article-level preference: neighbours' purchases rarely contain the query's next articles, so every
embedding row is below global popularity; the frozen TabPFN state tracks the aggregates it was fed (0.23 vs 0.26).
Repeat purchases dominate this task (PastVisit 2.2), which no customer-level embedding can express. Next: reference rows
that isolate the cause (user-kNN over the raw purchase matrix, item-kNN), a past-only+fill hybrid, and item-aware
embeddings (two-tower with article embeddings) before claiming anything about relational FMs on rel-hm.

## 2026-10-02 J7c (cluster job 50341): rel-hm with reference CF rows — the scoring is fine, the embeddings are the weak part
Official evaluator, MAP@12 x100, val / test. Same 74,575 / 67,144 queries, 365-day history, k = 50.
- kNN-CF over the raw purchase matrix (sparse cosine user-kNN, no model): **1.16 / 1.20** — 4x popularity, above the
  published LightGBM (0.38) and GraphSAGE (0.80) test rows. ItemKNN: 1.04 / 1.07.
- Past+kNN-CF[purchase_matrix]: 1.92 / **2.22** — the only hybrid above PastVisit alone (1.90 / 2.20).
- Our frozen-embedding rows unchanged: kNN-CF[row] 0.14, [agg] 0.26, [tabpfn_agg_kmeans] 0.23, [tabpfn_agg_random] 0.17,
  [tabpfn_row_kmeans] 0.14 (test); their Past+ hybrids 2.16-2.19, all slightly *below* PastVisit (the filler pushes popular
  items out of the padded slots).
Reading: the same kNN-CF machinery gets 1.2 when the neighbourhood is defined by co-purchase, and 0.14-0.26 when it is
defined by the customer table + aggregates (with or without the frozen TFM on top). The item-level signal is in the
interaction matrix, which the customer-level features (and therefore the TFM hidden state over them) do not carry.
Next (J7d): feed the interaction signal itself to the frozen TFM — truncated SVD(64) of the purchase matrix (+ aggregates)
as the TFM input, k-means / random target — and compare kNN-CF[svd], [svd_agg], [tabpfn_svd_kmeans], [tabpfn_svd_random]
against kNN-CF[purchase_matrix]. If the TFM row is below the raw SVD row, the "frozen TFM as graph-aware embedding" story
is dead at item level on rel-hm; if above, the TFM adds something over the factors.
Run: reports/runs/20261002T163801Z_J7_hm_full (cluster branch).

## 2026-10-02 J7d (cluster job 50399): the interaction signal through the frozen TFM — the hidden state is worse than its input
Official evaluator, MAP@12 x100, val / test, same queries, k = 50, TabPFN v2 get_embeddings (float32), k-means / random target.
- kNN-CF[svd] (64 SVD factors of the log1p purchase matrix, no model): 0.82 / **0.83** (vs raw sparse cosine 1.16 / 1.20).
- kNN-CF[svd_agg] (factors + customer aggregates): 0.66 / 0.65 — the aggregates dilute the interaction signal.
- kNN-CF[tabpfn_svd_kmeans] (frozen TabPFN over svd_agg, k-means target): 0.40 / **0.42**; random target 0.51 / 0.48.
- All Past+ hybrids 2.18-2.20, none above PastVisit (2.20) or Past+kNN-CF[purchase_matrix] (2.22).
Reading: given the interaction factors as input, the frozen TFM's hidden state keeps roughly half of their retrieval
value (0.83 -> 0.42-0.48); the k-means pseudo-target, which helped on the synthetic segments, does not help here (below the
random target). On rel-hm at item level the frozen-TFM-as-graph-embedding hypothesis is closed: every embedding row is
below the model-free factors, which are themselves below plain sparse cosine on the purchase matrix.
Run: reports/runs/20261002T173110Z_J7d_hm_svd_full (cluster branch).

## 2026-10-02 J12 (cluster job 50618): segment-level probe on rel-hm user-churn — frozen embedding = raw aggregates, no gain
20,000 customers labelled at the last train timestamp (2020-08-31, churn rate 0.817), 365-day history, 5-fold over customers,
kNN probe (k = 50, cosine, mean neighbour label), AUROC:
- Supervised HGB on the aggregates **0.673**; kNN[agg] 0.653; kNN[tabpfn_agg_kmeans] 0.648; kNN[tabpfn_agg_random] 0.643;
  kNN[tabpfn_svd_kmeans] 0.645; kNN[svd_agg] 0.638; kNN[svd] 0.590; kNN[row] 0.520; prior 0.5.
- The supervised-TabPFN reference row failed (inference_precision must be a torch dtype) -> fixed, rerun as J12b.
Reading: on a real entity-level label the frozen TabPFN hidden state neither destroys nor adds information relative to its
input (0.648 vs 0.653 for the same kNN on the raw aggregates), and both sit 2 AUROC points below a supervised GBDT on the
same features. The k-means pseudo-target gives +0.6 over a random target here (synthetic: +0.3 P@10). So the segment-level
version of H1 holds only in the weak form "a frozen TFM embedding is as good as the hand aggregates it was fed"; it is not a
better entity representation than its input on rel-hm. (Not comparable to the official user-churn leaderboard: random
customer folds at one timestamp, not the temporal test split.)
Run: reports/runs/20261002T203451Z_J12_churn_full (cluster branch).

Addendum (J12c, job 50659): the supervised-TabPFN reference row is 0.672 AUROC, equal to HGB (0.673). So the same frozen
model extracts the label when the label is in its context, and its label-free embedding + kNN (0.648) is 2.4 points below
that: the price of the pseudo-target is exactly the gap between "frozen embedding" and "frozen model used as a classifier".

## 2026-10-02 J14 (cluster job 50826): Kumo Relational on real data — entity level (rel-hm user-churn kNN probe, AUROC)
Same 20,000 customers / 5 folds as J12. Frozen Kumo Relational (sdm) over the rel-hm graph customers <- transactions ->
articles (365-day window, <= 200 transactions per customer, 2 hops, 64 context rows), readout-token state, 172 s on one H200.
- kNN[kumo_relational_random] **0.660**; kNN[kumo_relational_kmeans] 0.644; kNN[agg] 0.653; kNN[tabpfn_agg_kmeans] 0.648;
  supervised HGB 0.673, supervised TabPFN 0.672.
Reading: the relational model's graph embedding with a *random* in-context target is the best training-free row, 0.7 AUROC
points above the hand aggregates and 1.2 above the tabular TFM over them; it is the first row where a frozen embedding adds
anything beyond its tabular input, and the gain is small (still 1.3 points below supervised). Here the k-means target
hurts: it pulls the embedding toward the aggregate clusters and discards the graph signal, the opposite of the synthetic
result (0.33 random -> 0.62 k-means), where the planted segments *were* the aggregate clusters. Item-level rows (user-item
-purchase, official evaluator) follow in the same job.
Run: reports/runs/20261002T231936Z_J14_kumo_churn (cluster branch).

## 2026-10-03 J14 item level (job 50826): Kumo Relational on rel-hm user-item-purchase, official evaluator, MAP@12 x100
- kNN-CF[kumo_relational_kmeans] 0.20 / **0.18** (val / test); hand aggregates 0.26; sparse cosine on the purchase matrix
  1.20; PastVisit 2.20; Past+kNN-CF[kumo] 2.18 (no gain over PastVisit alone).
Reading: at item level the relational hidden state is no better than the tabular one: like every embedding row, it is
below the popularity prior. The small entity-level gain (churn 0.660 with a random target) does not translate into
item-level neighbourhoods; what the graph embedding carries is a customer-level summary, not co-purchase structure. The
random-target variant at item level is queued (J14b) for completeness; the churn ordering (random > k-means) suggests it
will be higher than 0.18 but the gap to 1.20 is two orders of magnitude.
Run: reports/runs/*J14_kumo_hm_full (cluster branch).

## 2026-10-03 J14b (job 50921): Kumo Relational, random target, item level — 0.25 / 0.24 (val / test MAP@12 x100)
Random target 0.24 vs k-means 0.18 (same ordering as the churn probe), level with the aggregates (0.26), 5x below sparse
cosine on the purchase matrix (1.20); Past+kNN-CF 2.18 < PastVisit 2.20. Retrieval study complete: no frozen embedding,
tabular or relational, with any in-context target, is an item-level retrieval representation on rel-hm.

## 2026-10-03 J20 (TEmBed, IBM 2026): TabPFN context target on a public row-embedding benchmark — no effect on text-heavy rows
Wikidata-books row triplets (anchor closer to positive than negative; chance 0.50), 5 variants, mean accuracy:
MiniLM text embeddings 0.77; TabPFN all-zeros target (TEmBed's default) last layer 0.58, layer 8 0.58; random 0.55;
k-means last 0.55, layer 8 0.56. The k-means target does not help here: the rows are titles / authors / genres, which
TabPFN sees as arbitrary category codes. Entity-matching row similarity could not be run (TEmBed's preparation script
fails with a missing-config error). Conclusion: the context-target effect, if real, is about numeric structure; text
tables need a text encoder.
