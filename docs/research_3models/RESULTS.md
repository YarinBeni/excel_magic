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
| pi F: + labels + verified ledger | 0.210 | 0.161 | 5.2 | 5.8 | see paper | -0.063 (0.030) |

pi D vs our loop D: -0.032 (0.025). Reading: the tool server works unchanged under a second harness; the result
pattern is the same as in our loop (DEEP tools neutral to negative, the verified ledger costs recall). pi's tool calls
were not logged.
