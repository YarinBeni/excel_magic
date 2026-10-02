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
