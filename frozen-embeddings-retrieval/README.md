# frozen-embeddings-retrieval

**Research question:** are the hidden states of *frozen* foundation models for structured data (tabular FMs such as
TabPFN, relational FMs such as KumoRFM-style models) useful **graph-aware entity embeddings** for retrieval, with no
training on the target database? How do they compare with row-level embeddings, hand aggregates, and a GNN trained
on the database?

Companion library: [open-tabfm-auto](https://github.com/YarinBeni/open-tabfm-auto) (backbones, weights, run logging).

## Setup

```bash
pip install -e ".[dev]"             # pulls open-tabfm-auto[tabpfn] from GitHub
tabfm-models download tabpfn        # TabPFN v2 weights (GCS mirror)
curl -L -o weights/openrfm/openrfm-pretrain-ctx96.pt \
  https://github.com/T-Lab/OpenRFM/releases/download/v0.1.0/openrfm-pretrain-ctx96.pt
python experiments/exp01_retrieval_benchmark.py                       # synthetic shop DB, all embedders
python experiments/exp01_retrieval_benchmark.py --kind northwind --db northwind.db
```

## Embedders (`fer/embedders.py`)

| name | what it is | trained on the DB? |
|---|---|---|
| `row` | standardized customer-table columns | no |
| `agg` | hand-written relational aggregates (counts, recency, spend, category mix) | no |
| `tabpfn_row` / `tabpfn_agg` | TabPFN v2 per-row hidden state over the row / aggregate table, random target (no label signal) | no (frozen) |
| `tabpfn_agg_churn` | same, fitted with a real downstream label | no (frozen) |
| `openrfm` / `openrfm_ctx16` / `openrfm_random` | 256-d pre-readout state of OpenRFM (KumoRFM-2 reproduction), customer row + last order lines + tickets as root/child/aux; random-weight control | no (frozen) |
| `gnn` | HeteroGraphSAGE trained with customer-product link prediction, customer hidden state | **yes** |
| `kumo_relational` | NVIDIA KumoRelational readout tokens (forward hook) | no (frozen); needs HF weights |

## Benchmarks (`fer/retrieval_bench.py`)

- **A. latent-segment retrieval** (synthetic DB): each customer has a hidden segment that only influences *which
  products* they buy, so it is invisible in the customer row and recoverable only through relations. P@10, MAP@10,
  10-NN accuracy.
- **B. future-purchase retrieval** (any DB with orders): kNN collaborative filtering in embedding space, MAP@10 /
  recall@10 against post-cutoff purchases, popularity baseline.

## Headline result (synthetic shop DB, 2000 customers, 6 hidden segments, 3 seeds)

Segment retrieval P@10 (chance 0.167) and 10-NN segment accuracy; future-purchase kNN-CF MAP@10 (popularity 0.034).

| embedder | trained on DB? | seg P@10 | seg 10-NN acc | fut MAP@10 |
|---|---|---|---|---|
| row (customer columns) | no | 0.161 | 0.146 | 0.022 |
| agg (hand relational aggregates) | no | 0.726 | 0.842 | 0.048 |
| tabpfn_row | frozen | 0.169 +/- 0.005 | 0.171 | 0.021 |
| tabpfn_agg, one random target | frozen | 0.492 +/- 0.344 | 0.567 | 0.036 |
| tabpfn_agg, 4 random targets averaged | frozen | 0.525 +/- 0.325 | 0.617 | 0.037 |
| **tabpfn_agg, k-means pseudo-labels** | frozen | **0.824 +/- 0.021** | **0.873** | **0.048** |
| tabpfn_agg, 4 held-out feature targets | frozen | 0.680 +/- 0.030 | 0.787 | 0.042 |
| gnn (HeteroGraphSAGE link prediction) | **yes** | 0.877 +/- 0.007 | 0.902 | 0.047 |
| openrfm (KumoRFM-2 reproduction) | frozen | 0.249 | 0.306 | 0.032 |
| openrfm, random weights (control) | - | 0.168 | 0.169 | 0.013 |

What this says so far:
- The hidden state of a frozen tabular FM over a flattened relational table *can* match a GNN trained on the database
  (0.82-0.88 vs 0.88), but only with a structure-preserving in-context target: a single random target is a coin flip
  (0.21 to 0.88 across seeds); k-means pseudo-labels make it stable and beat the raw aggregates it was fed (0.73).
- The in-context target is a free "query conditioning" knob: fitting with a downstream label collapses the embedding onto
  that label (0.18 for churn).
- The open relational-FM reproduction carries real but weak relational signal (0.25 vs 0.17 random); the real test needs
  KumoRelational / Relational Transformer weights (Hugging Face).

Every number: `docs/run-log.md` (mean +/- sd over seeds 0-2, `scripts/aggregate_runs.py`).

## Real benchmark: RelBench rel-hm (official evaluator, test MAP@12 x100)

GlobalPopularity 0.29, PastVisit 2.20, published ID-GNN 2.81; **our frozen-embedding kNN rows 0.14-0.26, below popularity; classic user-kNN on the raw purchase matrix 1.20 and Past+kNN-CF 2.22 (> PastVisit 2.20), so the kNN scoring is fine and the customer-level embeddings are what lack the item signal; feeding SVD factors of the purchase matrix through the frozen TabPFN keeps only half of their value (0.83 raw factors -> 0.42 TabPFN), so at item level on rel-hm the frozen hidden state is worse than its own input; at entity level (user-churn, kNN probe AUROC) it equals its input (0.648 vs 0.653 raw aggregates; supervised HGB 0.673).**
Customer-level embeddings over coarse aggregates do not carry article-level preference (105k articles); see
`docs/run-log.md`. This is the honest state of the real-data test; the synthetic-DB results above do not transfer yet.

See `docs/PAPER_PLAN.md` for the hypotheses, the experiment matrix and what is still blocked.
