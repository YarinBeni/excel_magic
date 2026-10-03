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

**2026-10-02 J12 status note.** Segment level on real data (rel-hm user-churn, kNN probe): frozen TabPFN over the customer
aggregates 0.648 AUROC vs 0.653 for the raw aggregates and 0.673 for supervised HGB. The embedding is as good as its input,
not better. Combined with J7c/J7d, the defensible paper claim is narrow: frozen TFM hidden states with a k-means
in-context target are a *target-agnostic* entity representation that preserves (synthetic: recovers) the structure of the
features it is given, at no training cost; they do not add relational signal the features lack, and they lose to the
interaction matrix for item retrieval. A paper needs either a setting where that target-agnosticity is the point (many
downstream tasks per entity, one embedding) or a relational FM whose hidden state actually carries neighbour information
(OpenRFM 0.25 and Kumo Relational 0.62 did not beat TabPFN-over-aggregates 0.82 on the synthetic segments).

**2026-10-02 J14 status note.** Kumo Relational on rel-hm user-churn: 0.660 AUROC by kNN with a random in-context target,
above the aggregates (0.653) and the tabular TFM over them (0.648), below supervised (0.672). First real-data evidence that
a *relational* frozen hidden state carries graph signal the hand features lack, by a small margin; the k-means target
(0.644) removes it. If the item-level rows (pending) show the same ordering, the paper's positive claim becomes: frozen
relational-FM hidden states are weak but real graph-aware entity embeddings, and the in-context target must be
uninformative (random) to keep that signal.

**2026-10-03 J14 item-level note.** Kumo Relational (k-means target) 0.18 test MAP@12 x100 on user-item-purchase, below
the aggregates (0.26) and the purchase-matrix cosine (1.20). The relational FM's small entity-level gain does not reach
item level. Final framing for the paper: frozen FM hidden states (tabular or relational) are entity-level summaries;
the relational model adds a little graph signal at entity level (0.660 vs 0.653) and nothing at item level; neither is
a retrieval embedding for recommendation, where the interaction matrix itself is the right representation.

**2026-10-03 J14b, final.** Random-target Kumo Relational at item level: 0.24 (k-means 0.18, aggregates 0.26, purchase
matrix 1.20). The retrieval study is complete; the item-level claim is closed for both model families, the entity-level
gain of the relational model (0.660 vs 0.653) stands as the single small positive.

## Layer study (J18 / J19 / J21, started 2026-10-03): which layer is the entity embedding?

**Gap.** Layer-wise work on tabular FMs exists, but it covers *prediction*, not *label-free entity embeddings*:
"Is One Layer Enough?" (ICML 2026) measures CKA / cosine redundancy across the layers of six TFMs; "A Mechanistic Study of
Tabular Foundation Models" (May 2026) finds TabPFN v2's target probe jumps between layers 8 and 9 while TabICLv2 is readable
early; TEmBed (IBM 2026) benchmarks only the last-layer `get_embeddings`. No one has asked, on real relational benchmarks,
(a) which layer of a frozen FM gives the best entity embedding, (b) how the in-context target changes that, (c) whether a
label-free statistic can pick the layer, and (d) whether a *relational* FM's layers behave like a tabular FM's.

**Hypotheses.**
- **H5 (depth)** An inner layer beats the last layer as an embedding; the last layers specialise to the in-context target.
- **H6 (context)** With real labels in context, late layers become label read-outs; with random or k-means context the
  embedding stays general. The best layer moves with the context.
- **H7 (label-free selection)** A geometric statistic of the layer output (effective rank, TwoNN intrinsic dimension,
  anisotropy, k-means silhouette) picks a layer close to the validation-picked one, with no labels.
- **H8 (convergence)** Different models agree more at their best layers than at their last layers (CKA).

**Protocol.** 8 RelBench entity classification tasks (rel-f1 driver-dnf / driver-top3, rel-trial study-outcome, rel-event
user-repeat / user-ignore, rel-avito user-visits / user-clicks, rel-hm user-churn), official splits and evaluator. Each
task row becomes a temporally safe 2-hop subgraph (Kumo Relational) or its flattened relational features (tabular FMs).
Probe = logistic regression on the frozen layer output, trained on 4000 train rows disjoint from the context; layer picked
on validation; test AUROC reported; kNN probe as a second view. Baselines: probe on the raw features; the model's own
prediction (real-label context). Models: Kumo Relational (J18), TabPFN v2 with four context targets (J19), TabPFN v2,
TabPFN-2.5, TabICLv2, Kumo Tabular S/M/L with random / k-means / real-label context plus geometry and CKA (J21; 5 tasks).
Aggregator: `scripts/analyze_layers.py` -> `docs/layers/LAYERS.md` + 4 figures.

**Preliminary (rel-f1 only, 2 tasks; not yet a result).** Kumo Relational, real-label context: the graph layer (before the
in-context transformer) is val-picked on both tasks, 0.812 / 0.858 test vs its own prediction 0.783 / 0.877 and its last
layer 0.787 / 0.771. TabPFN v2 over the flattened features: k-means context 0.833 mean val-picked vs all-zeros last layer
0.693; no layer beats the raw-feature probe on average (0.842).

**Protocol correction (2026-10-03, J21 first pass).** These first-pass numbers used a random split of context vs
probe-train rows, which leaks under a real-label context: probe-train rows had their own entity in context at a
neighbouring date, so late layers carried a copied label that val/test rows never get. Six tabular FMs then showed
below-chance late-layer linear probes (TabICLv2 0.40, Kumo-S 0.43) while kNN on the same layers held 0.75-0.79. The
real-label "inner layer beats last layer" evidence above is therefore biased toward early layers. All studies rerun
with a time split (probe-train rows strictly later than every context row). The leak itself is worth a paragraph in the
paper: layer-probing studies of in-context models must keep probe rows out of the context's period, or late layers look
uninformative when they are in fact specialised to the context.
