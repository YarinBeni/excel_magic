# Reproducing the paper's TabArena experiment

The paper (Sec. 4.1, App. B.1) runs, per dataset: one pipeline search on the training split of fold 0 (3-fold inner
CV, up to 96 evaluations or 6 h on one H100), then scores the frozen P* on every official split (9 or 30 per dataset,
816 in total), and rates each configuration by Bradley-Terry Elo against the 66-method TabArena pool.
`tabfm_auto/benchmarks/tabarena.py` implements exactly that protocol:

```bash
pip install -e ".[tabpfn]"           # `openml` is a core dependency
tabfm-models download tabpfn         # or kumo-tabular-s / tabicl once Hugging Face is reachable

# smoke test: 3 smallest datasets, Lite (split r0f0 only), tiny budget  (~30 min CPU, ~$1 with Sonnet)
python examples/05_tabarena_protocol.py --max-instances 1000 --lite --budget-evals 6 --budget-minutes 15

# the 17 datasets with <= 2.5k rows, all official splits               (~1 day CPU or ~2 h on a GPU node)
python examples/05_tabarena_protocol.py --max-instances 2500 --budget-evals 24 --budget-minutes 60

# paper-scale: 51 datasets, one SLURM array task each, 96 evals / 6 h, Opus
sbatch scripts/slurm_tabarena.sh
```

Outputs in `runs/<id>/`: `results.csv` (dataset, repeat, fold, method, error per split), `comparison_to_paper.csv/.md`
(our P0 / P* per dataset next to the paper's TabFM and TabFM-Auto numbers from Table 9, with relative gains), one
search directory per dataset with the agent stream and every candidate pipeline.

## Elo against the leaderboard pool

`tabfm_auto.benchmarks.elo.elo_ratings(df, anchor="RandomForest (default)", n_bootstrap=100)` reproduces the
paper's rating (Bradley-Terry, RF default = 1000, bootstrap-over-tasks interval) from a long frame of
`(method, task, error)`. The pool's per-split errors are published by TabArena (the `tabarena` package downloads them
from Hugging Face / the TabArena cache); concatenate them with our `results.csv` rows and call `elo_ratings`.
Without the pool, `comparison_to_paper.md` gives the per-dataset comparison against Table 9 directly.

## What differs from the paper, and why it matters

| paper | here | effect |
|---|---|---|
| TabFM 400M (Google, closed) | TabPFN v2 (11M) today; Kumo Tabular-S / TabICLv2 / EXAONE via `--model` once weights are reachable | the frozen model's own strength sets the baseline; the *gain* from the pipeline is the comparable quantity |
| Claude Code + Opus 5, 96 evals / 6 h per dataset | any Claude model via `--llm`, any budget | budget drives search depth; Sonnet stops early when gains flatten |
| bubblewrap sandbox | subprocess with network stripped, test rows outside the workspace | best-effort isolation |
| context up to 16,384 rows | `--max-rows` (TabPFN v2 context <= 10k on CPU; raise on GPU) | large datasets rely on `sample()` views |
| 5 configurations x 51 datasets ~ $17.6K LLM spend | your choice | estimate: Sonnet ~$0.2-2 per dataset at 24 evals; Opus ~5-10x |

## Compute estimates (TabPFN v2, n_estimators=8)

| datasets | rows | one 3-fold eval on 4-core CPU | on one H100 |
|---|---|---|---|
| 17 smallest (<= 2.5k rows) | 748-2400 | 10-60 s | 1-3 s |
| mid (2.5k-13k rows, e.g. churn, heloc, jm1) | 3k-13k | 1-6 min | 5-20 s |
| large (> 20k rows, context capped by `--max-rows`) | 20k-150k | 3-15 min | 20-60 s |
| wide (Bioresponse 1776 cols, hiva 1617 cols) | | P0 fails (> 500 columns) as in the paper; the agent must select/compress columns | |
