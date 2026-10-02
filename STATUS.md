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
- J7c done (50341): user-kNN on the raw purchase matrix 1.20 test MAP@12 x100 (4x popularity, > published LightGBM/GraphSAGE),
  Past+kNN-CF[purchase_matrix] 2.22 > PastVisit 2.20; embedding rows unchanged 0.14-0.26. The kNN step works; the customer-level
  embeddings lack the item signal. inbox 017 / J7d: SVD(64) of the purchase matrix (+agg) through the frozen TabPFN vs raw factors.
- J0v (50374) OK: vllm 0.16.0, torch 2.9.1+cu128. J3 (50375) / J5 (50376) / J6 (50377) released.
- J5 task 2 (GLM-4.5-Air, 50406) failed at vLLM engine start: 106B-A12B bf16 (~220 GB) does not fit one H200 -> GLM-4.5-Air-FP8
  at 0.90 GPU memory; inbox 018 resubmits array task 2 only. vLLM failures now print the engine-side root-cause lines.
- J5 task 3 (gpt-oss-20b, 50376) failed: openai_harmony tiktoken cache under /tmp not writable (another user's dir) ->
  TIKTOKEN_RS_CACHE_DIR=~/.cache/tiktoken-rs in vllm_env; inbox 019 resubmits array task 3 only.
- J6 qwen-code (50408): every turn 400 from vLLM: Qwen Code requests max_tokens = 32768 = the whole server context, leaving 0
  input tokens. vLLM now serves 128k by default (VLLM_MAX_LEN; Qwen3-32B 40k). J6 task 1 resubmitted as 50428 (picks the
  fix up at job start); J3/J5/J6 jobs already running keep their 32k server (our openai harness sets its own max_tokens).
- J3 (50375) + J5 task 0 (50402), Qwen3-Coder-30B via our openai harness: tools used correctly, but gains are small or negative
  (physics +1.3/+2.9%, entities -1.1/+1.0%, breast_cancer -2.9/-56.7%, TabArena-Lite -0.9..+2.5%); heuristic finds more on
  synth_entities. Harness: run_eval on an unchanged pipeline now refused (one run looped 38x). RESULTS.md / REPORT 5.2 updated.
- J5 task 3 (gpt-oss-20b, 50425): vLLM OK, but read_file on a .parquet raised UnicodeDecodeError and killed every search ->
  tool errors go back to the model; inbox 021 resubmits task 3.
- J6 task 2 aider + Qwen3-Coder (50377): synth_entities +41.7% (0.384 -> 0.224; frequency encodings + 4 context views),
  physics -2.2%, breast_cancer -40%. Same LLM in our tool loop never found the encodings -> harness matters more than the LLM here.
- J6 task 1 Qwen Code + Qwen3-Coder (50428, 128k context): physics +1.0%, entities +25.5%, breast_cancer -16.1%. Entity ranking
  with the same LLM: aider +41.7% > Qwen Code +25.5% ~ heuristic > our tool loop ~0.
- J5 task 3 gpt-oss-20b (50431): physics -1.3%, entities -9.7%, breast_cancer +7.9%; harmony parser leaked channel markers
  into tool names (now normalised); TabArena run died on choices=None (now guarded; run errors keep a traceback). inbox 022 reruns.
- J7d done (50399): SVD factors of the purchase matrix 0.83 test MAP@12 x100 by plain kNN; frozen TabPFN over them 0.42 (k-means)
  / 0.48 (random). The frozen hidden state is worse than its input -> H1 closed at item level on rel-hm (segment-level only).
- J5 task 3 gpt-oss-20b rerun (50434, clean harness): physics +4.1%, entities -8.9%, breast_cancer -7.9%, TabArena-Lite +2.0/-1.1/-0.2%.
- J5 task 2 GLM-4.5-Air-FP8 (50414): physics +5.1%, entities +43.6% (best of all runs), breast_cancer -38.9%, TabArena-Lite
  -0.9/+1.4/+1.4%. Best open LLM in our tool loop. Remaining: Qwen3-32B (J5 task 1), pi (J6 task 0).
- J6 task 0 pi (50377_0) failed in 4 min: conda create node22 raced with task 1 (same env); creation now under flock, inbox 024
  resubmits task 0. Qwen3-32B (J5 task 1) still running.
- J6 task 0 pi + Qwen3-Coder (50515): entities +50.3% (best of all), physics -0.9%, breast_cancer -17.1%. Only Qwen3-32B left.
- J5 task 1 Qwen3-32B (50405): physics -0.2% (1 eval), entities -3.8%, breast_cancer -2.9%, TabArena-Lite +0.7/+0.6/+0.3%. Weakest.
  LLM sweep (J3/J5) and CLI sweep (J6) complete.
- inbox 025 / J8: TabArena wave 1 (17 datasets, all official splits) with pi + Qwen3-Coder (task 0) and tool loop + GLM-4.5-Air-FP8
  (task 1), 16 evals / 40 min per dataset, to compare with J2 (heuristic +0.5%) and the paper (+4.6%). Results land per dataset.
- J8 task 1 (GLM, 50554_1) failed at vLLM start: 128k KV (23 GiB) does not fit next to the FP8 weights (18 GiB free) -> 64k;
  inbox 026 resubmits task 1. J8 task 0 (pi) running: 4/17 datasets done.

## 2026-10-02 next steps (inbox 027)
- Library: repeated k-fold judge (`--cv-repeats N|auto`), rich first message for the minimal tool loop (TABFM_LOOP_RICH=1).
  Research: exp03 rel-hm user-churn segment-level probe.
- J9 per-backbone TabArena (heuristic x Kumo-L / TabICLv2), J10 budget 64 + auto repeats on anneal/Marketing/airfoil/hazelnut/
  diabetes (heuristic, pi, GLM), J11 cv3 judge on < 1000-row tables (heuristic, pi), J12 churn probe, J13 rich loop (Qwen3-Coder).
- Goal: end with the best open-source option (backbone x harness x LLM x judge) with evidence, in README/REPORT.
- J8 pi + Qwen3-Coder done 17/17: mean -0.9% (median 0.0%, 8/17 wins) vs heuristic +0.5% vs paper +4.6%; loses where the
  paper gains most (anneal -5.2%, Marketing -14.2%). GLM loop 6/17 so far (-2.9%). REPORT 5.3 / RESULTS updated.
- J12 done (50618): rel-hm user-churn kNN probe AUROC: HGB 0.673, kNN[agg] 0.653, kNN[tabpfn_agg_kmeans] 0.648 -> frozen
  embedding = its input, not better. Supervised-TabPFN row failed (dtype arg) -> fixed, inbox 028 reruns (J12b).
- J13 done (50619): rich-context tool loop with Qwen3-Coder: physics -0.7%, entities -4.9%, breast_cancer -23.5%, Lite +3.3/+0.4/+1.4%.
  Does not close the gap to pi -> harness effect is the agent loop itself. Early J9: Kumo-L (4 ds) -2.8%, TabICLv2 (6 ds) -9.2%;
  J10 budget64 heuristic (3 ds) +3.3%, pi (5) -4.3%; J11 cv3: heuristic entities +8.3%, pi entities +28.1% / breast -5.4%.
- J12c (50659): supervised TabPFN on agg 0.672 = HGB; frozen embedding + kNN 0.648 -> the pseudo-target costs 2.4 AUROC points.
- J11 (cv3 judge) heuristic done, pi 4/5: on the <1000-row TabArena tables heuristic +1.4% -> +2.1%, pi -1.4% -> -0.5%;
  breast_cancer pi -17% -> -5%. Recommend --cv-repeats auto. REPORT 6 / RESULTS updated.
