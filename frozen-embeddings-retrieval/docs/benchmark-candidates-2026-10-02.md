# Published benchmarks to compare against (research note, 2026-10-02)

Produced by a web-research subagent; arxiv / openreview / HF were egress-blocked, so numbers marked (snippet) come from
search snippets and must be re-verified against the PDFs before being quoted in a paper.

## Ranked candidates

### 1. RelBench link-prediction (recommendation) tasks  <- primary target
Robinson et al., NeurIPS 2024, https://arxiv.org/abs/2407.20060 ; code https://github.com/snap-stanford/relbench ;
data on Hugging Face `stanford-star/relbench`; leaderboard https://star-project.stanford.edu/relbench/leaderboard/.
Task: temporal link prediction; for each source entity at time t rank destinations it links to in (t, t+delta].
Metric MAP@K (rel-hm user-item-purchase K=12, 7-day window; rel-amazon K=10; others vary). Fixed timestamp splits.

Published test MAP x100 (RelBench Table 5, snippet-derived):

| task | GlobalPop | PastVisit | LightGBM | RDL-GraphSAGE | RDL-ID-GNN |
|---|---|---|---|---|---|
| rel-amazon user-item-purchase | 0.24 | ~0 | 0.16 | 0.74 | 0.10 |
| rel-amazon user-item-rate | 0.15 | 0.07 | 0.17 | 0.87 | 0.12 |
| rel-amazon user-item-review | 0.11 | 0.04 | 0.09 | 0.47 | 0.09 |
| rel-hm user-item-purchase | 0.30 | 0.89 | 0.38 | 0.80 | 2.81 |
| rel-avito user-ad-visit | ? | ? | 1.95? | ? | 3.66 |
| rel-stack user-post-comment | 0.02 | 1.42 | 0.04 | 0.11 | 12.72 |
| rel-stack post-post-related | 1.46 | 1.74 | 2.00 | ? | 10.83 |
| rel-trial condition-sponsor-run | 2.52 | 8.42 | 4.82 | ? | 11.36 |
| rel-trial site-sponsor-run | 3.75 | 17.31 | 8.40 | ? | 19.00 |

Later results (test MAP x100, snippet): ContextGNN (ICLR 2025, https://arxiv.org/abs/2411.19513): amazon purchase 2.93,
rate 2.25, review 1.63, hm 2.93, stack user-post-comment 13.34, trial condition-sponsor-run 11.65, site-sponsor-run 28.02;
8-task averages ContextGNN 9.23, NBFNet 7.71, GraphSAGE 2.08, LightGBM 2.01. KumoRFM v1 zero-shot rel-hm 2.73 (OpenRFM
Table 13, https://arxiv.org/abs/2606.04320; OpenRFM + agentic featuriser 2.56).
Compute: baselines trained on one A6000; an H100 is ample. Popularity / PastVisit / LightGBM are CPU-cheap.
Our slot-in: embed users and items with frozen hidden states using data up to the split timestamp; score items by
two-tower cosine or user-kNN CF; evaluate MAP@K with `task.evaluate`. Honest comparison rows: PastVisit, GlobalPop,
LightGBM, KumoRFM zero-shot; ID-GNN / ContextGNN are the supervised ceiling.

### 2. TEmBed row-similarity search (entity-matching retrieval)  <- secondary target
"Towards Universal Tabular Embeddings: A Benchmark Across Data Tasks", https://arxiv.org/abs/2604.21696 ;
code https://github.com/IBM/table-representation-evals. Top-k row similarity search on 9 entity-matching datasets
(Abt-Buy, Amazon-Google, DBLP-ACM, ...), metric MAP; TFM rows embedded via `get_embeddings()` - exactly our
frozen-hidden-state setting. Published mean MAP (snippet): GritLM 0.76, Granite-R2 0.73, MiniLM 0.66, TabuLa-8B 0.44,
HyTrel 0.05, TabICL v2 0.04, SAP-RPT-1 0.03. Frozen TFM row embeddings are a weak baseline there; a relational
variant (rows enriched by joins) would be a novel extension.

### 3. 4DBInfer key prediction (Wang et al., NeurIPS 2024 D&B, https://arxiv.org/abs/2404.18209)
Link tasks Diginetica purchase / Amazon-Book purchase / MAG cite, metric MRR; R-GAT and HGT numbers published; data at
HF `stanford-star/dbinfer`. Good secondary benchmark.

### 4. Methodological neighbours (cite / compare, not benchmarks)
OpenRFM (https://arxiv.org/abs/2606.04320), RDBLearn (https://github.com/HKUSHXLab/rdblearn), RelICL
(https://github.com/uma-pi1/relicl): frozen TabPFN/TabICL + parameter-free relational encoders.

### 5-7. TabEmbedBench (single-table, no retrieval), RelArena (no link tasks), OGB (not DB-like): skip or supplementary.

## First experiment (when Hugging Face is reachable)
RelBench `rel-hm` / `user-item-purchase` (MAP@12). Reproduce GlobalPopularity (0.30), PastVisit (0.89), LightGBM (0.38) on
CPU; run the official ID-GNN script once on an H100 (target 2.81). Add our rows: (a) frozen-TabPFN customer / article
embeddings + two-tower cosine; (b) user-kNN CF over those embeddings; (c) hybrid PastVisit + kNN-CF fill. Report next to
KumoRFM zero-shot 2.73 and ContextGNN 2.93. Second task: `rel-amazon` / `user-item-purchase` (MAP@10).
