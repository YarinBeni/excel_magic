# Status board (newest first)

## 2026-10-02 branch created
- Queued via inbox: 001 (J0 setup -> J1 smoke), 002 (J4 RelBench probe), 003 (J2 TabArena heuristic x17), 004 (J3 vLLM open-LLM search).
- Waiting on: Yarin to clone the branch on the cluster, set the push token, start the runner.
- Expected first results: `reports/J0_models_status.md` (which backbones loaded), `reports/runs/*J1_*` (backbone comparison,
  heuristic search with Kumo Tabular-S, retrieval incl. KumoRelational), `reports/J4_relbench_probe.md`.

## 2026-10-02 (later) added
- inbox 005: J5 LLM sweep (Qwen3-Coder-30B-A3B, Qwen3-32B, GLM-4.5-Air, gpt-oss-20b via vLLM, one GPU each, after J0).
- library: `--harness cli --agent-cmd <pi|qwen-code|gemini-cli|aider|codex|template>` for free coding-agent CLIs (J6 to follow once
  the headless flags are verified).

## 2026-10-02 J0 done (job 50193, exit 0, H200): weights for TabPFN v2/2.5, TabICLv2, Kumo Tabular-S, Kumo Relational, EXAONE downloaded;
   tests 15+2 passed. Gap: NVIDIA sdm not installed (pip name is structured-data-models) -> inbox 007 re-runs J0 and resubmits J1-J6.
- J4 probe OK: rel-hm loads (15.4M tx); val MAP@12 x100: GlobalPop 0.34, PastVisit 1.90. -> inbox 008 / J7: full rel-hm experiment with our rows.
- J1 (partial): Kumo-S best open backbone on synth_physics (0.087) / synth_entities (0.447); TabICLv2 best on wine/diabetes. Bugs fixed:
  --harness choices in the TabArena example, comma-in-spec model lists. inbox 009 resubmits J2.
- J1 (partial): Kumo-S best open backbone on synth_physics (0.087) / synth_entities (0.447); TabICLv2 best on wine/diabetes. Bugs fixed:
  --harness choices in the TabArena example, comma-in-spec model lists. inbox 009 resubmits J2.
- J3/J5/J6 died racing on the vLLM venv creation -> _vllm.sh installs under flock with an .ok marker; J0v installs once; inbox 010 resubmits.
- J1 done (50213): Kumo Relational hook works (512-d; seg P@10 0.48 / 0.26 over seeds, random target); TabPFN embeddings failed on GPU
  (bf16) -> fixed (float32); heuristic+Kumo-S: synth_entities -29% (0.462->0.328). J7 smoke failed on subset evaluation -> fixed.
  inbox 011 reruns J1 and J7.
- J2 wave 1 complete (17/17): mean gain +0.5% (paper +4.6%), wins 9/17, Kumo-S P0 within a few % of TabFM on 16/17. J0v failed on
  libstdc++ (venv on thesis python) -> vLLM in its own conda env; inbox 012 reruns J0v and re-chains J3/J5/J6.
- J7 done (50266, official evaluator, test MAP@12 x100): GlobalPop 0.29, PastVisit 2.20, our embedding kNN rows 0.14-0.26 (below
  popularity), hybrids 2.19 (no gain). NEGATIVE for the frozen-embedding hypothesis at item level; recorded in fer run-log/README/
  PAPER_PLAN and REPORT 5.6. inbox 014 reruns J7 with reference rows (user-kNN on the purchase matrix, item-kNN). J0v (50316)
  rerunning with the LD_LIBRARY_PATH fix; J3/J5/J6 chained behind it.
- J0v (50316) failed on `vllm --version` ("Failed to infer device type": the CLI parser needs a GPU, J0v is a CPU job); the
  libstdc++ problem is gone. Check is now a plain import; inbox 015 reruns J0v and re-chains J3/J5/J6.
- J0v (50342) installed vllm 0.30.0 = torch 2.13+cu130; the node driver is CUDA 12.8 -> cuda unavailable, every vLLM server
  in J3 (50343) / J5 (50362, 50365) died at engine start. Installer now pins vllm 0.16.0 (torch 2.9.1 cu128; fallback 0.11.0)
  and asserts a CUDA 12 torch build at install time. inbox 016 reruns J0v and re-chains J3/J5/J6.
