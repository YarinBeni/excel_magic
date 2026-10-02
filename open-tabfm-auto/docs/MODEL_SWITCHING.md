# Switching backbones

Everything that depends on a model goes through one string, the model spec, and one directory, `weights/`.

## 1. See what is runnable

```bash
tabfm-models status          # table: model, state (ready / needs huggingface.co / downloadable), Elo, licence, CPU notes
```

## 2. Get weights (the moment huggingface.co is allowed, or on a machine that has it)

```bash
tabfm-models download kumo-tabular-s tabicl tabpfn-2.5 exaone kumo-relational   # or: tabfm-models download all
rsync -a other-machine:open-tabfm-auto/weights/ weights/                                 # side-load from the cluster instead
```

All Hugging Face based libraries are pointed at `weights/hf-cache` (`HF_HOME`), TabPFN at `weights/tabpfn`,
so the whole state of "which models can run" is that one folder. It is git-ignored.

## 3. Use a different model

| what | how |
|---|---|
| any experiment | `--model kumo-tabular-s:n_estimators=4` (exp01/02/03), `--models a,b,c` (exp01/05) |
| from code | `tabfm_auto.models.get_model("tabicl", task_type)` |
| the agent's pipeline search | the spec is written into the workspace `task.json`, the agent never sees model code |
| relational embeddings | `--embedders openrfm,kumo_relational` in exp04 |

Available spec names: `tabpfn` (v2, mirror, runnable now), `tabpfn-2.5`, `tabpfn-2.6`, `tabpfn-3`, `tabpfn-3.5`,
`tabpfn-3.5-fast`, `tabicl`, `tabiclv2-sdm`, `kumo-tabular-s|m|l`, `exaone`, and the classical `hgb`, `rf`,
`lightgbm`, `logreg`, `dummy`. Key=value options after `:` go to the constructor (`n_estimators`, `device`, ...).

## 4. Compare backbones without re-running the search

```bash
python examples/04_backbone_transfer.py --models tabpfn,kumo-tabular-s,tabicl,hgb
```

re-scores the already discovered pipelines (P0 and P*) of every search run on the same held-out split with each
runnable model, i.e. the paper's transfer experiment (Figure 5). Expected order of strength on TabArena:
kumo-tabular-l (~1960) > kumo-tabular-m (~1910) > tabpfn-3.5 (~1855) > kumo-tabular-s (~1790) > exaone (1765) >
tabpfn-3 (1660) > tabpfn-2.6 (1605) > tabicl (1590) > tabpfn-2.5 (1522) > tabpfn v2.

## 5. Adding a new model

1. Add a `ModelCard` to `tabfm_auto/models/manifest.py` (repo id, files, mirror, licence, Elo).
2. Add a builder in `tabfm_auto/models/registry.py` returning an object with `fit`, `predict`, `predict_proba`
   (see `sdm_wrapper.py` for wrapping a non-sklearn API).
3. `tabfm-models status` should show it; `exp05` picks it up automatically.
