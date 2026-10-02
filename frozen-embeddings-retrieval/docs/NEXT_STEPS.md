# Next steps once Hugging Face and the GPU cluster are available

1. `pip install relbench` and fetch `rel-hm` (HF `stanford-star/relbench`); `tabfm-models download kumo-relational tabicl kumo-tabular-s`.
2. Reproduce the published rel-hm / user-item-purchase rows on CPU: GlobalPopularity, PastVisit, LightGBM
   (RelBench repo `examples/`), and ID-GNN once on an H100 (official script). Verify the numbers in
   `docs/benchmark-candidates-2026-10-02.md` against the PDFs.
3. Our rows via `fer/relbench_adapter.py` -> `ShopDB` -> `fer.embedders` (tabpfn_agg_kmeans, kumo_relational, openrfm,
   gnn) -> two-tower cosine and user-kNN CF -> `task.evaluate(..., "test")` MAP@12.
4. TEmBed row-similarity benchmark (IBM table-representation-evals): plug `tabpfn_agg_kmeans`-style embeddings in as an
   approach; compare to the published GritLM / TabICL v2 rows.
5. Paper table: 3 seeds, cost column (seconds per 1k entities), ablation on the in-context target (random / k-means /
   feature / downstream label).
