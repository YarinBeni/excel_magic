# How it works (and how it maps to the paper)

TabFM-Auto (Fu et al., 2026, arXiv 2609.37989) keeps a tabular foundation model (TFM) **frozen** and lets an LLM
coding agent evolve the *data pipeline* around it. Because nothing is trained, every candidate pipeline is cheap to
evaluate and the search signal is clean. This repository implements the same loop with open-weights TFMs.

```
          +-----------------------------+        tabfm-eval         +------------------------------+
          |  Coding agent (Claude Code) | ----------------------->  |  Judge (3-fold CV, sandboxed) |
          |  edits workspace/pipeline.py| <-----------------------  |  frozen TFM, train split only |
          +-----------------------------+   score + traceback       +------------------------------+
                        |                                                      |
                        | best CV candidate frozen after the budget            | never sees
                        v                                                      v
               best_pipeline.py  ---- scored once ---->  held-out split (outside the workspace)
```

| Paper component | This repo |
|---|---|
| Pipeline `P = (preprocess, engineer, sample, postprocess)` + `TABFM_KWARGS` (Sec. 3.2) | `tabfm_auto/pipeline/template.py` (identity P0), contract enforced in `pipeline/runner.py` (<= 500 columns, context views, probability alignment) |
| Frozen TabFM 400M | any backbone in `tabfm_auto/models/manifest.py`: TabPFN v2 (runs today), TabPFN 2.5-3.5, TabICLv2, Kumo Tabular S/M/L, EXAONE-Tabular |
| Sandboxed judge, 3-fold CV on the training split, test labels unmounted (Appendix A) | `tabfm-eval` -> `harness/evaluator.py` in a subprocess with proxies stripped and a timeout; test split lives in `runs/<id>/_heldout/`, outside the agent's working directory |
| Agent harness: Claude Code / Codex / Antigravity | `agent/claude_code.py` (headless `claude -p`), `agent/openai_compat.py` (any OpenAI-compatible LLM, e.g. vLLM + Qwen3-Coder), `agent/heuristic.py` (no LLM: greedy search over generic operations); every event logged |
| Dataset metadata M (column names, task description, aux files) | `TabularTask.metadata` rendered into `workspace/TASK.md` |
| Search once per dataset, evaluate the frozen P* on held-out folds | `agent/search.py::run_search` writes `metrics.json` with `p0_cv`, `best_cv`, `p0_test`, `best_test`, classical baselines |
| Transfer of P* to other TFMs (Sec. B.3) | `examples/04_backbone_transfer.py` |
| (new) training table from a relational database | `agent/db_agent.py`: a first agent session explores the DB read-only with `tabfm-db`, writes `build_table.sql` (one row per entity, features before the cutoff, label after); the harness materialises it and runs the normal search |

## Workspace layout the agent sees

```
workspace/
  TASK.md          rules + dataset brief + P0 score
  pipeline.py      the only file it should edit
  train.parquet    features + target (training split only)
  task.json        task type, metric, model spec, budget, seed
  metadata.json    column descriptions
  evals.jsonl      one record per tabfm-eval call (score, folds, features, timing, errors)
  candidates/      snapshot of pipeline.py for every evaluation
```

## Isolation caveat

The paper uses a bubblewrap namespace sandbox. Here isolation is best-effort (environment stripped of proxy
variables, offline flags, timeout, tool allow-list, held-out rows outside the workspace). A malicious pipeline could
still read files elsewhere on disk; do not run untrusted agents on sensitive data without a container.
