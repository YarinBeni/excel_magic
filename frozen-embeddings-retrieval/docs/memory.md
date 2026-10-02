# Memory (append-only)

## 2026-10-02
- Repo split out of the tabfm-lab prototype. Depends on open-tabfm-auto for backbones / logging.
- Findings: H1 and H2 supported at seed 0; H3 untestable with OpenRFM (weak); need HF weights.
- Next: 3-seed table, Northwind, linear-probe task (D), kumo-tabular-s / tabicl row-embedding hooks once weights arrive,
  node2vec baseline, cost table.
- 3-seed check (2026-10-02): `tabpfn_agg` with ONE random binary target is unstable: seg P@10 0.875 / 0.210 / 0.392 over seeds 0/1/2
  (mean 0.49 +/- 0.34), while gnn is 0.877 +/- 0.007 and agg 0.726. H1 as first stated is NOT supported; the in-context
  target dominates the hidden state. Testing variance-reduced targets: average over 4 random targets (`tabpfn_agg_rand4`),
  k-means pseudo-labels (`tabpfn_agg_kmeans`), held-out feature columns as targets (`tabpfn_agg_feat4`).
- Northwind (jpwhite3 extended SQLite, 93 customers / 16k orders): every customer has bought nearly every product, so both the
  future-purchase and novel-purchase tasks saturate (MAP@10 = 1.0 / undefined). Dropped as a benchmark DB; a real DB needs
  sparse customer-product interactions (RelBench rel-hm / rel-amazon once reachable).
- Added `future_purchase_novel` (candidates and truth exclude products bought before the cutoff) and `scripts/rescore.py`
  (re-score saved `emb_*.npy` with the current benchmark code) + `scripts/aggregate_runs.py` (mean +/- sd over seeds).
- Cluster J1b: KumoRelational + k-means target = 0.62 +/- 0.05 seg P@10 (random target 0.33 +/- 0.13). H2 (target choice) now shown on
  two model families. Next: KumoRelational with the TabPFN-kmeans embedding's own labels, more readout tokens / hops, and the
  RelBench rel-hm rows (J7b).
