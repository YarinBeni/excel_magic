# Paper plan: frozen structured-data foundation models as entity embedders

## Claim candidates (to be confirmed or killed by the matrix below)

- **H1** The hidden state of a frozen tabular FM over a flattened relational feature table is a competitive
  graph-aware entity embedding (no training on the DB), on par with a GNN trained on the DB.
  Evidence (3 seeds): with k-means pseudo-labels as the in-context target, seg P@10 0.824 +/- 0.021 vs GNN 0.877 +/- 0.007
  vs raw aggregates 0.726; with one random target it is unstable (0.49 +/- 0.34). H1 holds conditionally on the target;
  the target-choice result (H2) is the more interesting finding.
- **H2** The embedding is controlled by the in-context target: a random target keeps feature structure, a real
  downstream label projects the embedding onto label-relevant directions (0.875 -> 0.182). This is a feature, not a
  bug: "query-conditioned embeddings for free".
- **H3** Relational FMs (KumoRFM-style) embed the entity *with its neighbourhood* and should beat flattened tables
  when the relational signal is not captured by hand aggregates. Open reproduction (OpenRFM) is too weak to test
  this (0.25); needs KumoRelational / Relational Transformer weights.
- **H4** Frozen embeddings transfer across databases without any fitting (zero-shot), unlike the GNN.
- **Status after rel-hm (2026-10-02):** H1/H3 do NOT hold at item-level retrieval on real data: all embedding kNN rows are
  below global popularity (0.14-0.26 vs 0.29 MAP@12 x100). The synthetic finding (segment retrieval) measures a coarse
  group signal; rel-hm needs article-level preference. The paper angle must either (a) move to entity-similarity tasks
  where group structure is the target (TEmBed-style, customer segmentation, entity matching) or (b) make the embeddings
  item-aware (two-tower with article embeddings; frozen FM over user x item-feature interactions).

## Experiment matrix

| axis | values |
|---|---|
| embedder | row, agg, tabpfn_{row,agg}, tabpfn_agg_{random,label}, kumo-tabular-s (hook), tabicl (hook), openrfm, kumo_relational, rt-j, gnn, node2vec |
| database | synthetic shop (planted segment; vary n_segments, noise, n_customers), Northwind (real, small), RelBench rel-trial / rel-avito / rel-hm (official MAP@K) when reachable |
| task | A segment retrieval, B future-purchase retrieval (MAP@K), C entity dedup / analogues (Northwind customers by country), D downstream linear probe (churn AUROC from frozen embedding) |
| seeds | 3 (0,1,2), report mean +/- std |
| controls | random-weight model, shuffled relations, aggregates-only |
| cost | embedding seconds per 1k entities on CPU |

## Published benchmark to compare against (decided 2026-10-02, see docs/benchmark-candidates-2026-10-02.md)

Primary: **RelBench recommendation tasks**, first `rel-hm / user-item-purchase` (MAP@12, 7-day window), then
`rel-amazon / user-item-purchase` (MAP@10). Published rows to reproduce: GlobalPopularity 0.30, PastVisit 0.89,
LightGBM 0.38, RDL-GraphSAGE 0.80, ID-GNN 2.81, KumoRFM zero-shot 2.73, ContextGNN 2.93 (rel-hm, MAP@12 x100; verify
against the PDFs). Our rows are training-free (frozen hidden states + kNN), so the fair comparison is against
PastVisit / popularity / LightGBM / KumoRFM zero-shot, with ID-GNN / ContextGNN as the supervised ceiling.
Secondary: **TEmBed row-similarity search** (entity-matching retrieval, MAP), where TFM `get_embeddings()` rows are
already a published (weak) baseline. `fer/relbench_adapter.py` maps rel-hm onto our `ShopDB` views.

## What is blocked and by what

- KumoRelational, Relational Transformer (RT-J), Kumo Tabular S, TabICLv2 weights: `huggingface.co` denied by the
  environment network policy. Code paths exist (`fer/embedders.py::embed_kumo_relational`, `tabfm_auto.models`).
- RelBench data: relbench 3.x downloads from Hugging Face. Fallback: Northwind (GitHub), synthetic.
- GPU: not needed for the current scale (2k entities); needed for RelBench-size databases.

## Writing targets

- Workshop-length first: 4 pages, Tables A/B on synthetic + Northwind + 1 RelBench task, 3 seeds, ablation on target
  choice (H2), cost table.
- Keep every number traceable: `runs/<id>/metrics.json` -> `docs/run-log.md` -> paper table.

**2026-10-02 J7c status note.** Reference rows settle the attribution: user-kNN on the raw purchase matrix reaches 1.20
test MAP@12 x100 (4x popularity, above published LightGBM/GraphSAGE), and Past+kNN-CF[purchase_matrix] 2.22 is the only
hybrid above PastVisit. The frozen-embedding rows (0.14-0.26) therefore fail because customer-level features do not
carry item-level co-purchase structure, not because of the retrieval step. J7d tests the remaining version of H1: the
frozen TFM over SVD factors of the interaction matrix vs the raw factors.

**2026-10-02 J7d status note.** H1 is closed at item level on rel-hm: with the interaction factors as input (SVD(64) of the
purchase matrix, test MAP@12 x100 0.83 by plain kNN), the frozen TabPFN hidden state scores 0.42 (k-means target) / 0.48
(random). The paper's honest framing is therefore: frozen TFM hidden states are usable *segment-level* entity embeddings
when the in-context target is chosen well (synthetic 0.82 vs GNN 0.88), and they are not item-level retrieval embeddings
on a real purchase graph, where they lose to their own input and to sparse cosine on the interaction matrix.
