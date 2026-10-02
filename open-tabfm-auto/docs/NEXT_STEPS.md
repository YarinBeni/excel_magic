# Next steps once Hugging Face and the GPU cluster are available

1. Weights (one-time, on a node with huggingface.co access):
   ```bash
   tabfm-models download kumo-tabular-s tabicl tabpfn-2.5 exaone        # into ./weights (rsync-able)
   tabfm-models status
   ```
2. Backbone transfer on the pipelines already found (no LLM needed):
   `python examples/04_backbone_transfer.py --models tabpfn,kumo-tabular-s,tabicl,exaone,hgb`
3. LLM for the search, pick one:
   - open weights on the cluster: `vllm serve Qwen/Qwen3-Coder-30B-A3B-Instruct --enable-auto-tool-choice --tool-call-parser hermes`
     then `--harness openai --llm Qwen/Qwen3-Coder-30B-A3B-Instruct --llm-base-url http://<node>:8000/v1`
   - no LLM: `--harness heuristic`
   - Claude Code: `--harness claude-code --llm opus`
4. Paper protocol (needs api.openml.org): smoke test first
   `python examples/05_tabarena_protocol.py --max-instances 1000 --lite --budget-evals 6 --budget-minutes 15`
   then the SLURM array `scripts/slurm_tabarena.sh` (env: MODEL, LLM, BUDGET_EVALS, BUDGET_MINUTES).
5. Compare: `runs/<id>/comparison_to_paper.md`; Elo with `tabfm_auto.benchmarks.elo` once the TabArena pool results are
   downloaded (the `tabarena` package fetches them).

## 2026-10-02 evening: after the LLM / CLI sweeps
- J8 (running): TabArena wave 1 with pi + Qwen3-Coder and the GLM-4.5-Air tool loop; fills REPORT 5.3 next to J2.
- Judge for small tables: repeated 3-fold CV (or a holdout of the training split) when n < 1000; every open setup
  overfit breast_cancer.
- Tool loop vs CLI gap: give the minimal loop the same affordances (a file tree, eval history, the data head in the
  system prompt) and re-run Qwen3-Coder; if the gap closes, the harness effect is prompt engineering, not agency.
- Per-backbone search on TabArena with Kumo-L / TabICLv2 (paper 5.4 conclusion).
- Retrieval: the segment-level story needs a real benchmark with entity-level labels (RelBench node classification
  tasks with kNN probing) before any paper claim; item-level is closed on rel-hm.
