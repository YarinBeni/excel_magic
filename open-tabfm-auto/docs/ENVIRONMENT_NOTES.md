# Blockers and environment limits (2026-10-02)

| Blocker | Effect | Workaround used | Ask |
|---|---|---|---|
| Network policy denies `huggingface.co` (and `cdn-lfs.huggingface.co`, `hf-mirror.com`, `cas-bridge.xethub.hf.co`) | No TabICLv2, TabPFN v2.5+/3, Kumo Tabular, EXAONE-Tabular, Relational Transformer, KumoRelational, Griffin weights; no RelBench data (relbench 3.x downloads from HF) | TabPFN v2 from the public GCS mirror; OpenRFM from GitHub Releases | add `huggingface.co`, `cdn-lfs.huggingface.co`, `cas-bridge.xethub.hf.co` to the environment's allowed domains |
| `api.openml.org` / `www.openml.org` denied | No TabArena datasets | bundled sklearn + synthetic datasets | allow `api.openml.org`, `www.openml.org` |
| `relbench.stanford.edu`, `zenodo.org`, `archive.ics.uci.edu`, `data.pyg.org`, `snap.stanford.edu` denied | No RelBench / UCI / PyG datasets | synthetic shop DB with planted structure; Northwind SQLite from GitHub | allow as needed |
| `download.pytorch.org` denied | could not install the CPU-only torch wheel; the PyPI CUDA wheel (~5 GB) was installed instead | works on CPU, just large | allow `download.pytorch.org` |
| GitHub API blocked for repos outside the session scope | could not list OpenRFM releases through the API | parsed the README for the asset URL | none |
| `YarinBeni/yarin_claude_code` not visible to this session's GitHub connection (`list_repos` does not return it) | could not read the agentic-workflow conventions | mirrored the append-only run-log / memory style of `mentor-council/AGENTS.md` | grant the Claude GitHub App access to that repo, or paste its AGENTS/CLAUDE.md |
| CPU only (4 cores, 15 GB RAM), no GPU | TabPFN v2 ~5-60 s per 3-fold CV on 500-1500 rows; OpenRFM fine for 2k entities; large TFMs (LimiX-2 400M, Kumo-L 215M) impractical | small datasets, `n_estimators=2-4` | cluster access for TabArena-scale runs |
| No bubblewrap | pipeline-search sandbox is best-effort (network vars stripped, timeout); the agent could in principle `cat` files outside the workspace | held-out split kept outside the workspace, rules in the prompt, `--allowedTools` restricted | none for now |
