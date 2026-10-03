# Open tabular and relational FMs as a "fast statistician" tool for an LLM agent (research report, 2026-10-03)

Sourcing: ~25 searches; arXiv / HF / NVIDIA pages blocked, so exact limits were read from source code of PriorLabs/TabPFN,
tabpfn-extensions, soda-inria/tabicl, NVIDIA/structured-data-models and T-Lab/OpenRFM. [code] / [vendor] / [2nd-hand] /
[anecdote] mark the evidence. 2026 releases: TabPFN-2.6, 3 (May), 3.5 (Sep); Google TabFM (30 Jun); Kumo Relational
(NGC 18 Aug); Kumo Tabular (29 Sep).

## Hard limits
| model | max context rows | max features | classes | licence | notes |
|---|---|---|---|---|---|
| TabPFN v2 | 10K [code] | 500 | 10 | Prior Labs licence (Apache-2.0 + attribution), commercial OK | categorical if <30 unique; CPU refuses >1K rows by default |
| TabPFN-2.5 / 2.6 | 50K | 2,000 | 10 | non-commercial; login / token needed | ~8-16 GB GPU |
| TabPFN-3 / 3.5 | 1M (H100) | 20,000 | 10 (+many_class ECOC) | non-commercial | vendor: 20x faster than 2.5; dates only with TRANSFORM_DATES |
| TabICLv2 | ~100K best; 500-600K with offload | 2,000 | many with flag | BSD-3, commercial OK | ~28M params; quantile regression; KV cache; vendor 50K x 100 in <10 s on H100 |
| Kumo Tabular S/M/L | no hard cap (per-estimator context subsets) | first 500 (default recipe) | 10 (+ECOC) | OpenMDW 1.1 commercial [press] | 28M / 115M / 215M; 999 quantiles |
| Kumo Relational v1 | 20K contexts, 2 hops (benchmark) | numeric + datetime per table | 10 | NGC terms (verify) | text excluded in benchmark; ~1 min per task |
| KumoRFM-2 (closed) | 10K contexts (BEST), 6-hop sampling | - | - | SaaS | PQL interface, NL explanations |
| OpenRFM | small | - | - | MIT | ~69 vs 79.6 AUROC of KumoRFM-2 |
| Mitra | ~5K best | ~100 | - | Apache-2.0 | in AutoGluon 1.4 |
| SAP-RPT-1-OSS (ConTextTab) | 8,192 (2,048 CPU / API) | - | - | verify | native semantic text |
| CARTE | fine-tuned, best <2K | - | - | open | native text (fastText) |

## Accuracy
TabArena single-model Elo (2026, 2nd-hand): AutoGluon 1.5 extreme 1695, TabPFN-3 1673, TabPFN-2.6 1626, RealTabPFN-2.5
1604, TabICLv2 ~1597, TabDPT 1460, CatBoost 1418, XGBoost ~1375. Kumo Tabular press claim TabArena #1 (newer board, not
comparable). BeyondArena (2026, 142 datasets, 100-1M rows): TFMs dominate small-to-medium IID data; tuned GBDT / MLP
win non-IID (temporal, grouped), large, wide and text-rich slices; TabPFN-3.5's own report concedes temporal / grouped /
large. Hardware: on a T4, TabICL +0.8 pp on Higgs-100K at ~40,000x tree latency (960 s, 9 GB VRAM). Relational: KumoRFM-2
79.60 average AUROC on 12 RelBench tasks (+1.54 over RelGNN); open Kumo Relational RelArena Elo 1832 (rank 3/12), ~1 min
per task vs LightGBM 14 min, GraphSAGE 47 min.

## Scaling past the context
Subsample / bag per estimator (TabPFN SUBSAMPLE_SAMPLES; sdm per-estimator row blocks); cache the context and batch the
query side (fit_with_cache, kv_cache); kNN retrieval context (LoCalPFN, TabDPT); TabICL offload; fine-tuning; TabPFN
distillation to MLP / trees (sub-ms CPU); practitioners stratify-sample, push joins and aggregates to SQL, fall back to
GBDT above ~100K rows or on temporal splits.

## Interpretability and uncertainty
tabpfn-extensions: SHAP via shapiq (imputation-based reuses KV cache; TabPFN-3 claims 120x faster SHAP), partial
dependence / ICE, feature selection, pval_crt (conditional-randomisation-test p-values: the most "statistician-like"
tool), decoder readout (prediction as weighted vote of training rows), embeddings, outlier scores, imputation,
synthetic data; regression returns full predictive distributions. TabICLv2 / Kumo: quantiles; no explanation API in
sdm (use permutation importance / SHAP with cached fit). Drift (suggested): classifier two-sample test, old vs new rows,
AUROC = drift score, SHAP = attribution.

## Text
TabPFN local: high-cardinality strings ordinal-encoded ("usually adds noise"); TRANSFORM_TEXT gives 30 tf-idf+SVD
features per column; hosted API embeds text. Kumo Tabular recipes: TF-IDF / SentenceTransformer. Kumo Relational
benchmark excluded text. Evidence (Mraz et al., ICML 2025 workshop): text embeddings beat dropping text on 11/13
datasets; no single best embedder; reduce to ~16-64 dims. BeyondArena: GBDT / MLP still win text-rich slices.

## Practitioner experience
OOM fixes (subsample, test batches of 1K-10K, memory_saving_mode); CPU unusable above 1K-5K rows; deterministic with
fixed seed on the same hardware; IDs and high-cardinality codes ordinal-encoded; dates need manual features; 10-class
cap; default temperature 0.9 sharpens probabilities. Reddit anecdotes: "fast and low-effort, usually slightly behind
tuned XGBoost", "heavy to run inference", one "garbage values on 10K x 200". Licence friction: commercial self-hosting
options are TabPFN v2, TabICLv2, Mitra, Kumo Tabular (OpenMDW).

## Inside agents
Prior Labs TabPFN MCP server (fit_and_predict, upload_dataset -> data bypasses the LLM context) and a Claude Skill;
tabicl-mcp; DuckDB anofox_tabfm extension (FM from SQL); KumoRFM-2 "for humans and agents" (PQL + NL explanations);
CAAFE (LLM proposes features, TabPFN scores: ROC AUC 0.798 -> 0.822, 11/14 datasets); MLZero / AutoGluon-Assistant.
Recommended tool pattern: LLM writes SQL with a time cutoff -> tool samples N context rows, runs the FM with KV cache
under CV -> returns metrics, SHAP, PD, CRT p-values as JSON -> if n > 100K or temporal, also fit LightGBM and report both
-> multi-table: Kumo Relational instead of hand flattening.

## Sources
https://arxiv.org/abs/2511.08667 ; https://arxiv.org/abs/2605.13986 ; https://github.com/PriorLabs/TabPFN ;
https://github.com/PriorLabs/tabpfn-extensions ; https://docs.priorlabs.ai/troubleshooting/OOM-errors ;
https://docs.priorlabs.ai/agentic/mcp ; https://pypi.org/project/tabicl-mcp/ ; https://arxiv.org/abs/2602.11139 ;
https://github.com/soda-inria/tabicl ; https://github.com/NVIDIA/structured-data-models ;
https://build.nvidia.com/nvidia/kumo-relational/modelcard ; https://arxiv.org/abs/2604.12596 ;
https://kumo.ai/docs/rfm/configuration ; https://github.com/PriorLabs/relarena ; https://github.com/T-Lab/OpenRFM ;
https://mindfulmodeler.substack.com/p/the-state-of-tabular-foundation-models ; https://huggingface.co/papers/2606.30410 ;
https://arxiv.org/abs/2512.00888 ; https://huggingface.co/autogluon/mitra-classifier ; https://arxiv.org/abs/2410.18164 ;
https://arxiv.org/abs/2506.10707 ; https://arxiv.org/abs/2402.16785 ;
https://research.google/blog/introducing-tabfm-a-zero-shot-foundation-model-for-tabular-data ;
https://arxiv.org/abs/2507.07829 ; https://arxiv.org/abs/2406.05207 ; https://arxiv.org/abs/2305.03403 ;
https://arxiv.org/abs/2505.13941 ; https://duckdb.org/community_extensions/extensions/anofox_tabfm
Gaps: GitHub issue threads and HF discussions not readable; Kumo Tabular numbers and licence from press only; TabDPT,
SAP-RPT-1-OSS and Kumo Relational NGC licences to verify.
