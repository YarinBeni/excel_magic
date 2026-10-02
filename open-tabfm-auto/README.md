# open-tabfm-auto

**Self-evolving data pipelines around frozen tabular foundation models.** An open-source implementation of
[TabFM-Auto](https://arxiv.org/abs/2609.37989) (Google Research, 2026).

- **Frozen model, evolving pipeline**: an LLM coding agent writes the cleaning, feature engineering, context selection and
  calibration code around TabPFN / TabICLv2 / Kumo Tabular; nothing is trained, every candidate is scored by sandboxed 3-fold CV.
- **Any LLM, or none**: Claude Code, any OpenAI-compatible endpoint (vLLM + Qwen3-Coder on your own GPU), or a no-LLM heuristic search.
- **Tables or databases**: start from a CSV, or point the agent at a SQL database and let it build the training table.
- **Paper protocol built in**: TabArena's 51 datasets and official splits, the paper's per-dataset numbers, Bradley-Terry Elo.

```
 your table / SQL DB  -->  [agent edits pipeline.py]  -->  frozen TFM  -->  3-fold CV score  -->  repeat
                                   ^                                              |
                                   +----------------------------------------------+
```

The paper's insight: tabular foundation models (TFMs) are strong zero-shot predictors but blind to column semantics.
An LLM reads the column names and task description and writes *code* (domain formulas, cleaning rules, entity
counts, context selection, calibration) around the frozen model. Nothing is trained, so every candidate is cheap to
evaluate and the search signal is clean. On TabArena this took the top five positions (2013 Elo vs 1785 for the bare
model). This repo reproduces the mechanism with open weights and the Claude Code CLI as the agent.

## Install

```bash
pip install "open-tabfm-auto[tabpfn,baselines] @ git+https://github.com/YarinBeni/open-tabfm-auto.git"
tabfm-models download tabpfn          # 73 MB from a public GCS mirror (no Hugging Face account needed)
# coding agent: either Claude Code ...
npm install -g @anthropic-ai/claude-code
# ... or any OpenAI-compatible endpoint (vLLM + Qwen3-Coder on your GPU), or no LLM at all (--harness heuristic)
```

## 60-second demo

```bash
tabfm-auto search --dataset synth_physics --budget-evals 10 --budget-minutes 20
```

```
P0 cv rmse = 0.1077                       # identity pipeline around frozen TabPFN
eval #2: rmse = 0.08511 NEW BEST          # agent added frequency*chord/velocity and its log
eval #3: rmse = 0.08321 NEW BEST          # dropped raw columns
...
{'p0_test': 0.1108, 'best_test': 0.0830, 'test_improvement_pct': 25.1, 'n_evals': 4}
```

Your own data:

```bash
tabfm-auto search --csv loans.csv --target defaulted --description "Consumer loans; predict default within 12 months"
```

From a database (the agent explores the schema read-only, writes `build_table.sql`, then searches):

```bash
tabfm-auto db --db shop.sqlite --target churned --id-col customer_id --cutoff 2025-07-01 \
  --question "Predict whether each customer places no paid order in the 90 days after the cutoff"
```

Every run writes `runs/<timestamp>_<name>/` with `config.json`, `events.jsonl`, `run.log`, `metrics.json`,
`best_pipeline.py`, the agent's full event stream and every candidate it tried. `tabfm-auto runs` tabulates them.

## What you get

| | |
|---|---|
| **Paper-faithful pipeline contract** | `preprocess / engineer / sample / postprocess` + model kwargs, enforced (<= 500 columns, context views, probability alignment) |
| **Sandboxed judge** | `tabfm-eval`: 3-fold CV on the training split only, subprocess with network stripped, timeout; held-out rows never enter the agent's workspace |
| **Pluggable frozen backbones** | TabPFN v2 (today), TabPFN 2.5 / 2.6 / 3 / 3.5, TabICLv2, Kumo Tabular S/M/L, EXAONE-Tabular. `tabfm-models status` shows what is runnable; `tabfm-models download <name>` fetches weights |
| **Agent harness** | three interchangeable: `claude-code` (the paper's), `openai` (any OpenAI-compatible endpoint: vLLM-served Qwen3-Coder / GLM on your own GPUs), `heuristic` (no LLM at all); `none` evaluates P0 only. See `docs/LLM_BACKENDS.md` |
| **SQL mode** | `tabfm-db` read-only exploration tool; leakage rules (features before cutoff, label after) in the prompt and checked at materialisation |
| **Transfer** | `examples/04_backbone_transfer.py` re-scores discovered pipelines with every backbone (paper Sec. B.3) |

## Results so far

| experiment | metric | identity P0 | agent P* | gain |
|---|---|---|---|---|
| synthetic physics regression (500 rows) | RMSE | 0.1108 | 0.0830 | -25% in 4 evals, $0.15 |
| SQL churn task, agent-built 38-feature table | 1-AUROC | 0.2930 | 0.2752 | -6%, beats hand-written table (0.3165) |

See `docs/RESULTS.md`. TabArena-scale evaluation needs OpenML access and a GPU-class budget; the harness supports
`--dataset openml:<name>`.

## Reproducing the paper

`tabfm_auto/benchmarks/tabarena.py` implements the paper's TabArena protocol (search on fold 0's training split,
score the frozen pipeline on all 816 official splits, Bradley-Terry Elo) and ships the paper's per-dataset Table 9 for
direct comparison. `examples/05_tabarena_protocol.py` runs it; `scripts/slurm_tabarena.sh` fans it out on a GPU
cluster. See `docs/REPRODUCING_THE_PAPER.md` for budgets and compute estimates.

## Docs

- `docs/HOW_IT_WORKS.md`: component-by-component mapping to the paper, workspace layout, isolation caveats
- `docs/MODEL_SWITCHING.md`: backbones, weights, licences, expected TabArena Elo
- `docs/LLM_BACKENDS.md`: running the search with Claude Code, an open-weights LLM on your cluster, or no LLM
- `docs/RESULTS.md`: every number with its run directory
- `docs/REPRODUCING_THE_PAPER.md`: TabArena protocol, Elo, compute and cost estimates
- `CONTRIBUTING.md`

## LLM providers

| `--harness` | what it uses |
|---|---|
| `claude-code` (default) | the `claude` CLI, `--llm sonnet\|opus` |
| `openai` | `--llm <model> --llm-base-url http://host:8000/v1` (vLLM, SGLang, Ollama, OpenAI, OpenRouter) |
| `heuristic` | no LLM: greedy search over generic operations |
| `none` | identity pipeline only |

## Citation

This is an independent reimplementation; please cite the original paper (`CITATION.cff`):

```bibtex
@article{fu2026tabfmauto,
  title={TabFM-Auto: Self-Evolving Pipelines for Tabular Foundation Models},
  author={Fu, Deqing and Su, Huangyuan and Sen, Rajat and Narayan, Taman and Sanghavi, Sujay and Das, Abhimanyu and Kong, Weihao},
  journal={arXiv preprint arXiv:2609.37989},
  year={2026}
}
```
Model weights keep their own licences (TabPFN: Prior Labs licence; Kumo Tabular: OpenMDW; EXAONE: non-commercial).
