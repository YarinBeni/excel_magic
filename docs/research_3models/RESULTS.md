# Verified, Fast, Deep: results log

Every number below is read from the run directories under `reports/runs/` on the cluster branch.

## P0 verifier study (V1, job 51837): BIRD Arcwise-Plat, 497 questions, 4,473 SQL candidates
Generator Qwen3-Coder-30B (1 greedy + 8 sampled); judge GLM-4.5-Air (other family). Greedy accuracy 0.692; oracle 0.746.

| signal | AUROC | within-question AUROC | best-of-N | s per item |
|---|---|---|---|---|
| query executes | 0.556 | 0.631 | 0.706 | 0 |
| generator logprob | 0.579 | 0.512 | 0.692 | 0 |
| self-consistency | 0.651 | 0.721 | 0.706 | 0 |
| GLiClass large v3, zero-shot | 0.685 | 0.695 | 0.712 | 0.005 |
| NLI cross-encoder (DeBERTa-v3-large) | 0.656 | 0.733 | 0.720 | 0.005 |
| LLM judge | 0.786 | 0.748 | 0.714 | 0.064 (sequential) |
| stack, no-model signals | 0.676 | 0.725 | 0.706 | |
| stack + GLiClass | 0.725 | 0.766 | 0.712 | |
| stack + judge | 0.791 | 0.762 | 0.714 | |

Schema linking, recall of the gold SQL's columns in the top 10 (top 20): GLiClass 0.789 (0.877), lexical 0.706 (0.811),
bge reranker 0.687 (0.802), bge-small embedding 0.645 (0.774), GLiNER bi-encoder 0.350 (0.530).
Reading: GLiClass is the best cheap verifier (zero-shot, 5 ms) and adds +0.05 AUROC to the no-model stack; the LLM judge
is still best (12x slower). GLiClass is the best schema linker. Best-of-N headroom is small here (0.692 -> oracle 0.746).

## P1 text columns as features (V2, job 51825): 18 text+tabular datasets, official splits
Mean change vs numeric+categorical only (acc / r2), datasets better than base, and better than embeddings:

| model | GLiClass labels | embeddings (bge-small, PCA 16) | labels + embeddings | labels+emb beat emb |
|---|---|---|---|---|
| Kumo Tabular-L | +0.155 (17/18) | +0.184 (16/18) | +0.190 (16/18) | 15/18 |
| LightGBM | +0.128 (17/18) | +0.145 (16/18) | +0.159 (16/18) | 18/18 |
| TabPFN-2.5 | +0.126 (14/17) | +0.160 (16/17) | +0.159 (14/17) | 10/17 |

Reading: H1 (labels beat embeddings) does not hold; labels are complementary: labels + embeddings is best for Kumo-L and
LightGBM. Task-aware labels do not beat task-agnostic ones. 6 errors (TabPFN-2.5 on a >10-class dataset).

## P3 InsightBench (V3, jobs 51847 and 51908): 100 tables, analyst agent (Qwen3-Coder-30B), open judge (GLM-4.5-Air)

| config | g_eval | rouge1 | insights | rejected | vs D (se) |
|---|---|---|---|---|---|
| D: LLM + SQL | 0.304 | 0.245 | 5.6 | 0 | |
| E: + DEEP tools | 0.266 | 0.235 | 4.9 | 0 | -0.038 (0.018) |
| F: + text labels + verified ledger (ratio-percent fix) | 0.230 | 0.201 | 3.2 | 2.5 | -0.074 (0.022) |

Reading: negative. InsightBench's planted insights are descriptive (counts, shares, trends on 500-row tables) and the
score rewards recall; the DEEP tools cost steps without adding descriptive insights, and the verified ledger rejects
~2.5 insights per table that the agent does not replace. Verification trades recall for precision; this benchmark only
measures recall.

## P2 SQL harness ablation (V5, job 51863): BIRD Arcwise-Plat, 498 questions, Qwen3-Coder-30B

| config | accuracy | vs A (se) | schema chars | LLM calls |
|---|---|---|---|---|
| A: LLM alone, greedy, full schema | 0.693 | | 4,994 | 1 |
| B: + GLiClass schema linking (top 20 columns + keys) | 0.667 | -0.026 (0.016) | 2,254 | 1 |
| C: + GLiClass triage and one revision | 0.677 | -0.016 (0.016) | 2,254 | 1.03 |
| D: + 8 candidates, cheap verifier stack (2-fold) | 0.705 | +0.012 (0.016) | 2,254 | 2.3 |
| F: D + LLM judge on the uncertain band | 0.703 | +0.010 (0.015) | 2,254 | 2.3 (+1.5 judge) |
| SC: 8 candidates, majority result, no small models | 0.705 | +0.012 (0.009) | 4,994 | 2 |

Reading: with a strong generator on a corrected benchmark, the small models do not beat plain self-consistency.
Schema linking halves the prompt but loses 2.6 points (top-20 linking misses a needed column in 38% of questions);
triage recovers one point. All differences vs A are within about one standard error except B.

## P4 guarded auto-research over the SQL harness (V5, job 51972): BIRD Arcwise-Plat, 498 questions
Researcher GLM-4.5-Air proposes one knob change at a time; generator Qwen3-Coder-30B. Split fixed before the loop:
search 298 / accept 75 / held-out 125. Keep rule: paired search gain > 1 SE and accept gain >= 0. Budget 13 attempts
(the space ran out), 1 kept: `triage=True` (search +0.013, se 0.010; accept +0.013).

Held-out (scored once): initial (greedy, full schema) 0.680 -> final (+ GLiClass triage and one revision) 0.720,
paired +0.040 (se 0.018).

Rejected: self-consistency, n = 4/8/16, stack verifier (search -0.010 to +0.003), GLiClass linking (-0.020),
reranker linking (-0.067). Noise floor: link_k changes with linking off do not change the pipeline, yet moved search
accuracy by -0.007 (vLLM greedy decoding is not bit-reproducible under batching).
Reading: the guards worked as designed: 12 of 13 proposals were rejected, including three whose accept gain was
+0.027. The one kept change also helped on held-out. Caution: in P2 (all 498 questions, with linking) triage was
-0.016 vs A, so the triage gain depends on the configuration and is at most a few points.

## P3 InsightBench through the pi harness (V6, job 51914): same 100 tables, tools via `vfd` tool server
Same model (Qwen3-Coder-30B) and judge (GLM-4.5-Air); pi coding agent calls the tools as shell commands.

| config | g_eval | rouge1 | insights | rejected | tables with no insight | vs pi D (se) |
|---|---|---|---|---|---|---|
| pi D: SQL | 0.273 | 0.194 | 6.8 | 0 | 15 | |
| pi E: + DEEP tools | 0.269 | 0.191 | 6.7 | 0 | 17 | -0.004 (0.026) |
| pi F: + labels + verified ledger | 0.210 | 0.161 | 5.2 | 5.8 | 33 | -0.063 (0.030) |

pi D vs our loop D: -0.032 (0.025). Reading: the tool server works unchanged under a second harness; the result
pattern is the same as in our loop (DEEP tools neutral to negative, the verified ledger costs recall). pi's tool calls
were not logged.

## P3 RelBench prediction questions (V4, jobs 51838, 52039, 52052, 52074): 8 tasks x 4 configs x 2 episodes
Agent Qwen3-Coder-30B over DuckDB (database cut at the test time + train_labels + test_rows); official test AUROC.

| task | D: SQL | E: + Kumo Tabular-L | R: + Kumo Relational | ER: both |
|---|---|---|---|---|
| avito clicks | 0.584 | 0.661 | 0.624 | 0.627 |
| avito visits | 0.634 | 0.641 | 0.653 | 0.662 |
| event ignore | 0.500 | 0.568 | 0.877 | 0.840 |
| event repeat | 0.556 | 0.711 | 0.794 | 0.754 |
| f1 dnf | 0.506 | 0.792 | 0.713 | 0.744 |
| f1 top3 | 0.749 | 0.784 | 0.863 | 0.863 |
| hm churn | 0.595 | 0.681 | 0.674 | 0.674 |
| trial outcome | 0.491 | 0.591 | 0.690 | 0.690 |
| **mean** | **0.577** | **0.679** | **0.736** | **0.732** |

Fixes during the run (rows of the affected cells were removed and rerun): agent SQL is interrupted after 120 s (one
runaway join stalled the job); deep_fit_predict predicts test rows in chunks; DEEP predict caps its time hold-out at
max_context rows (rel-hm and avito ran out of GPU memory and the agent fell back to SQL).
Reading: the clearest win of the project. With SQL only, the agent is near chance on 4 of 8 tasks; the tabular model on
the agent's own features beats SQL on 8 of 8 tasks, and the relational model, which needs no feature SQL, is best on
average. ER does not beat R: the agent mostly submits the relational scores as they are.

# Round 2: harness engineering (no training; decisions on development splits only)

## BIRD failure analysis (config A, 498 questions, 153 wrong)
wrong values 82, wrong row count 26, query fails 16, empty result 12, wrong/extra/missing columns 17; 12 wrong answers
compare a literal that is not stored in the database. Deterministic checks can reach at most ~30-45 questions.

## Q1 harness v2 ladder (V7, job 52316): BIRD dev split (373 questions = P4 search + accept), Qwen3-Coder-30B
| config | accuracy | vs A (se) | gained / lost vs previous step |
|---|---|---|---|
| A plain | 0.686 | | |
| K1 + corrected column descriptions | 0.697 | +0.011 (0.015) | |
| K2 + data profile | 0.697 | +0.011 (0.016) | |
| K3 + rule card | 0.657 | -0.030 (0.022) | +20 / -35 |
| G + deterministic gates (fail, empty, all-NULL, literal not stored -> closest stored values), <=2 revisions | 0.737 | +0.051 (0.021) | +34 / -4 |
| GD + GLiClass doubt signal | 0.740 | +0.054 (0.022) | +7 / -6 |
Reading: the deterministic gates are the gain; the rule card hurts (drops asked columns, LIMIT 1 where ties matter);
the GLiClass doubt signal is noise. Round 2 removes the rule card (GN, GN8, SC8N) and repeats A/K2/GN on gpt-oss-120b
and GLM-4.5-Air.

## Q3 RelBench hypothesis loop, smoke (V8, job 52317): rel-f1 driver-dnf, 1 episode
5 hypotheses tested (recent DNF rate, overall DNF rate, races in 3 months, days since last race, ...), no leakage,
none passed the gate (best +0.0026, se 0.0024); system = FM = 0.748. Full run: job 52341.

## Q4 InsightBench analyst v2 (V3, job 52318): 100 tables, Qwen3-Coder-30B, GLM-4.5-Air judge
| config | g_eval | vs D (se) |
|---|---|---|
| D plain (rerun; first run 0.304) | 0.315 | |
| P + data profile + analysis checklist | 0.298 | -0.017 (0.020) |
| PC + GLiClass coverage signal | 0.280 | -0.035 (0.020) |
Reading: negative. Generic analysis kinds pull the agent away from the goal-specific questions the planted insights
answer. Next: a question-driven agenda (Q0: LLM's first 6 of 12 drafted questions; Q: GLiClass ranks the 12 by
relevance to the goal).

## Q1 round 2 (V7, job 52338): same dev split, no rule card
| config | accuracy | vs A (se) | LLM calls |
|---|---|---|---|
| SC8N descriptions + profile, 8 candidates, majority vote | 0.713 | +0.027 (0.017) | 2 (8 samples) |
| GN descriptions + profile + gates | 0.732 | +0.046 (0.018) | 1.14 |
| GN8 GN with 8 candidates, vote among candidates that pass the gates | 0.735 | +0.048 (0.018) | 3.24 |
Reading: gates with one candidate beat 8-way voting at about half the calls; extra candidates add little on top of
the gates (GN -> GN8: +5 / -4 questions); with gates on, the rule card no longer matters (G vs GN: +22 / -24).
Decision (dev only): final config GN. Held-out (125 questions) scored once for A, SC8N, GN, GN8: job 52376.

## Q5 GLiClass fine-tuned on DEV questions only (V9, job 52390); the LLM is unchanged. Tested on the 125 HELDOUT questions
| job | zero-shot GLiClass | fine-tuned GLiClass | LLM judge |
|---|---|---|---|
| column linker: recall of gold columns in top 10 | 0.610 | 0.955 | |
| column linker: all gold columns in top 20 | 0.36 | 0.93 | |
| SQL checker: AUROC over 1,125 candidates | 0.702 | 0.741 | 0.764 |
| best-of-9 accuracy (greedy 0.680, oracle 0.736) | 0.736 | 0.720 | 0.712 |
Training: checker 3,348 candidates of DEV questions; linker 1,119 (question, 40-label view) examples. Caveat: the held-out
questions use the same 11 databases, so the linker partly learns these schemas (the realistic company setting: train
on past queries over your own warehouse). Next: end-to-end harness with the tuned linker on held-out (GNL vs GNLz vs GN).

## Q5b zero-shot GLi-family comparison (V10, job 52389): all questions / candidates, corrected descriptions in labels
| model | checker AUROC | linker recall@10 | recall@20 |
|---|---|---|---|
| GLiClass large v3.0 | 0.643 | 0.641 | 0.770 |
| GLiNER2 large v1 | 0.655 | 0.730 | 0.829 |
| GLiNER2.5-Decide (340M) | 0.597 | 0.778 | 0.873 |
GLiNER2.5-base (needs protobuf) and GLiNER2.5-Decide-1B (needs transformers 5) did not load. All ~5.5 ms per candidate.

## Q2 harness gain on another LLM (V7, job 52339): gpt-oss-120b, BIRD dev split
A 0.681 -> K2 0.694 (+0.013, se 0.020) -> GN 0.727 (+0.046, se 0.019). Same gain as Qwen3-Coder-30B (+0.046).

## Q3 RelBench hypothesis loop, full (V8, job 52341): 8 tasks x 2 episodes, budget 10 tests, Qwen3-Coder-30B
| task | FM alone | FM + accepted LLM hypotheses | gain |
|---|---|---|---|
| f1 driver-top3 | 0.863 | 0.927 | +0.064 |
| f1 driver-dnf | 0.749 | 0.752 | +0.003 |
| trial study-outcome | 0.690 | 0.692 | +0.002 |
| avito user-clicks | 0.624 | 0.604 | -0.020 |
| avito visits, event repeat/ignore, hm churn | | nothing accepted | 0 |
| mean | 0.740 | 0.746 | +0.006 (3 better, 1 worse) |
145 hypotheses tested: 8 accepted, 102 rejected by the time-later gate, 28 constant, 7 caught leaking future data.
Accepted examples: "average qualifying position in the last 3 races", "recent top-3 frequency", "consecutive finishes in
the last 4 races", "number of sponsors of a study". The avito loss: features that passed on the validation period hurt
on the test period (drift).
