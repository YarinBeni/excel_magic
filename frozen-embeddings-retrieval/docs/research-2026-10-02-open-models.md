# Research note (2026-10-02): open tabular / relational foundation models reachable from a CPU box with Hugging Face blocked

Produced by a web-research subagent; items marked (unverified) could not be fetched directly because
huggingface.co, arxiv.org, docs.priorlabs.ai, relbench.stanford.edu, openml.org and grouplens.org were
blocked from the sandbox too.

## 1. Open tabular FMs that run on CPU

| Model | Params / ckpt | Weights hosted | License | CPU notes |
|---|---|---|---|---|
| **TabPFN v2** | ~11M; clf 29 MB, reg 44 MB | HF `Prior-Labs/TabPFN-v2-clf/-reg` **and GCS fallback `https://storage.googleapis.com/tabpfn-v2-model-files/05152025/`** (only v2 files there) | Prior Labs License (Apache-2.0 + attribution) | `tabpfn` pkg: CPU allowed <=1000 rows by default; raise via `ignore_pretraining_limits=True`. Max 10k rows / 500 feats. |
| TabPFN v2.5 / v2.6 | ~10.7M | HF only | non-commercial | same CPU default |
| TabPFN-3 / 3.5 | 53M clf / 58M reg | HF only | 3: non-commercial; 3.5: academic + eval | CPU default limit 5000 rows |
| **TabICLv2** | 28M | HF only `jingang/TabICL`; `model_path=` accepts a local file | permissive (BSD-3) | CPU supported, GPU recommended. https://github.com/soda-inria/tabicl |
| LimiX-2 / 16M / 2M | 400M / 16M / 2M | HF `stable-ai/*`; also ModelScope | LimiX-2 non-commercial | 400M impractical on 4 cores |
| EXAONE-Tabular | ~21M | HF only; `from_pretrained(weights=local_path)` | code BSD-3; weights NC | CPU supported but slow |
| TabDPT 1.3 / Turbo | not verified | HF only `Layer6/TabDPT` | Apache-2.0 | CPU not documented |
| Mitra v2 (AutoGluon) | ~76M | HF only | Apache-2.0 | CPU 12-63x slower |
| **Kumo Tabular** (NVIDIA, 2026-09-29) | S 28M, L 215M | HF `nvidia/Kumo-Tabular` | OpenMDW-1.1 | `sdm` lib, torch 2.7+, GPU-native; plain PyTorch ops. #1 on TabArena (Elo 1950). https://github.com/NVIDIA/structured-data-models |

TabFM-Auto paper: https://arxiv.org/abs/2609.37989 (project page https://deqingfu.github.io/tabfm-auto/).

## 2. Open relational FMs with public weights

- **KumoRelational** (NVIDIA `sdm`, Apr 2026): ~30M; weights on HF (repo id unverified); in-context fit/predict over `RelatedTables`; internal RowEmbedding -> InvariantGNN -> 12-layer ICL block with 4 readout tokens (512-dim graph embedding) but no public embed() (needs a hook). KumoRFM-2 proper is SDK/API only.
- **Relational Transformer (RT)**, ICLR 2026, 22M params; checkpoints HF `stanford-star/rt-j`; preprocessed RelBench on HF. flex_attention runs eager on CPU. https://github.com/layer6ai-labs/relational-transformer
- **OpenRFM** (T-Lab, Jun 2026): from-scratch KumoRFM-2 reproduction; **weights on GitHub Releases** (`openrfm-pretrain-ctx96.pt`, 224 MB, d=256, 3 layers, MIT); ~69 AUROC avg on 12 RelBench clf tasks; CUDA strongly recommended; no embedding API. https://github.com/T-Lab/OpenRFM/releases
- **Griffin** (ICML 2025): weights HF `yamboo/Griffin_models`; PyG. https://github.com/yanxwb/Griffin
- **RelArena / TabPFN-Rel** (Prior Labs, Aug 2026): framework over RelBench v1 entity tasks; needs TabPFN-3 (HF). https://github.com/PriorLabs/relarena
- RelGT, ContextGNN, Rel-LLM: code only, trained per task, no pretrained weights found.

## 3. Retrieval benchmarks

- **RelBench recommendation tasks (v1)**, metric `link_prediction_map` @K: rel-amazon user-item-* (K=10), rel-hm user-item-purchase (K=12), rel-avito user-ad-visit (K=12), rel-stack user-post-comment & post-post-related (K=100), rel-trial condition-sponsor-run & site-sponsor-run (K=10).
- Hosting: relbench 3.0.1 uses `huggingface_hub` on `stanford-star/relbench-v1`; local paths accepted. relbench <=1.1.0 downloaded from `https://relbench.stanford.edu/download/` (blocked here).
- Non-HF relational DBs: Northwind SQLite https://raw.githubusercontent.com/jpwhite3/northwind-SQLite3/main/dist/northwind.db; Chinook https://github.com/lerocha/chinook-database/releases; Hetionet https://github.com/hetio/hetionet; CTU repo (live MariaDB).

## 4. TabArena data

51 tasks stored as OpenML datasets/tasks; no GitHub/S3 mirror found. https://github.com/autogluon/tabarena

## Recommendation

- Tabular FM: **TabPFN v2** (reachable GCS mirror, permissive, CPU workable). If one HF file can be side-loaded, TabICLv2 is the better backbone and is in the paper's transfer set.
- Relational model: **OpenRFM** (weights on GitHub); hook its encoder for embeddings. With HF access: Relational Transformer RT-J or KumoRelational.
- Benchmark: RelBench rel-trial / rel-avito MAP@K need HF. CPU-only fallback: temporal link-prediction MAP@K on Northwind / Chinook SQLite, or a synthetic DB with planted structure.
