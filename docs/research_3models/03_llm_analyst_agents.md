# Open reasoning LLMs as SQL/BI analyst agents: strengths, limits, benchmarks (research report, 2026-10-03)

Sourcing: ~27 searches; direct fetches of arXiv / leaderboards blocked, numbers from search extracts. [unverified] =
single secondary source; [background] = model-card fact not re-checked. Re-check leaderboard tops before citing.

## Key findings
- Open models are close to closed on short analysis: Kimi K2 tops DSGym QRData-Verified (63.7 vs GPT-5 61.8) and
  DAEval-Verified (92.8 vs Sonnet 4.5 91.7); ReViSQL (Kimi K2.6 + RL on 2.5k hand-verified items) reaches human level
  (92.96%) on a corrected BIRD set.
- Long-horizon / interactive / enterprise work fails for all: BIRD-Interact GPT-5 8.7% (conversational) / 17%
  (agentic), open 4-11%; DABstep-hard <= 37%; LongDS-Bench best 48% with a ~47-point drop from early to late turns
  (52-69% of failures are state errors).
- Benchmarks are noisy: annotation errors in 52.8% of BIRD Mini-Dev and 62.8% of Spider 2.0-Snow (Jin et al., VLDB
  2026); corrected-set ranks correlate Spearman 0.32 with the official board.
- Context and verification matter more than the model: dbt semantic layer lifts frontier models from 84-90% to
  98-100%; multi-candidate SQL with execution checks +6-20 points at 5-30x the calls.

## Open models (single-GPU feasibility)
gpt-oss-120b (MXFP4 on one H100; tool-call quirks under vLLM; QRData-V 47.95); Qwen3/3.5 A3B models (fit easily;
vLLM bugs with tool calls inside <think>; Qwen3-Coder-480B BIRD-Interact 10.8 / 4.2); GLM-4.5-Air FP8 on one H200
(GLM-4.5 vendor: 90.6% tool-call success, BFCL-v3 77.8, TAU-bench retail 79.7); Kimi K2.x, GLM-5, DeepSeek V3.2,
Qwen3.5-397B need an 8-GPU node. Tool-call reliability depends on the serving stack: K2 Vendor Verifier saw <20%
parse success on an initial vLLM setup, 76% after fixes. Long context: all 18 models in Chroma's context-rot study
degrade; keep schema + results well under ~32-64K tokens.

## Benchmark table
| benchmark | size / data | ground truth / judge | best open | best closed / overall | caveats |
|---|---|---|---|---|---|
| BIRD (dev/test, Mini-Dev) | 12,751 pairs, 95 SQLite DBs | gold SQL, execution match | Arctic-Text2SQL-R1 (single model); ReViSQL+K2.6 92.96 on corrected set | ~82% test [unverified] | 52.8% Mini-Dev label errors |
| BIRD-Interact | 600 (Lite 300), Postgres + KB + user simulator | execution + tests | Qwen3-Coder-480B 10.8 / 4.2 | GPT-5 8.7 / 17.0 | simulator dependent |
| BIRD-Critic | 530 PG / 570 multi | test scripts | Bird-Fixer 14B 38.1 | o3-mini 38.9 | SQL repair only |
| Spider 2.0 (Lite/Snow/DBT) | enterprise warehouses, ~812 cols avg | execution / table match | ReFoRCE+DeepSeek-V3 38.0 (Snow) | 96.7 reported (Snow) | 62.8% Snow label errors |
| InsightBench | 100 business tables | planted insights, LLM judge | LLaMA-3-70B 0.51 | GPT-4o 0.60 | judge dependent, 2024 models |
| DAEval-Verified (from InfiAgent-DABench) | CSV QA | exact match | Kimi K2 92.8 | Sonnet 4.5 91.7 | near saturated |
| DSGym QRData-Verified | stat / causal QA | exact | Kimi K2 63.7 | GPT-5 61.8 | trajectories fine-tune Qwen3-4B 45 -> 59 |
| DABstep | ~450, payments CSV + docs | exact | Kimi K2 28.8 hard | Sonnet 4.5 37.0 hard | needs doc reading |
| KramaBench | 104 pipelines, data lake | expert answers | - | Claude 3.7 55.8 (62 with perfect retrieval) | retrieval bottleneck |
| BLADE | 12 scientific datasets | analysis decisions | - | F1 44.8; model-choice precision <35% | decisions, not answers |
| DiscoveryBench | 264 real + 903 synthetic | gold hypotheses, LLM judge | - | ~25-32 HMS | noisy |
| LongDS-Bench (2026) | 68 tasks, 2,225 turns | per-turn exact | - | 48.5 | long-horizon decay |
| AgenticDataBench (2026) | 15 domains | per skill | Codex + Kimi-K2.5 best | - | joins / alignment weak |

## Failure modes
Schema/semantic errors dominate (81.2% of 4,602 wrong queries; 27.6% of Spider 2.0 errors from schema linking; GPT-4o
0% on BEAVER enterprise warehouses); silent plausible-but-wrong aggregates; benchmark-to-production gap (vendor: 91%
public vs 21% on company data, undeclared joins); long-horizon state loss; confirmation bias and p-hacking
(counter-example prompting 42 -> 56% rule discovery); tool-call loops and parser bugs; run-to-run inconsistency
(Salesforce: ~50% single sample -> ~80% with 10 candidates + consistency scoring); context overflow on wide schemas and
inability to inspect large result sets; 10-30x cost of test-time scaling.

## Verification and self-improvement, measured
Self-consistency +5.8 on BIRD dev (CHASE-SQL; best-of-candidates ceiling +19.8); trained pairwise selector 73.0 dev;
Agentar-Scale-SQL 81.67 test; ReFoRCE (execution feedback, column exploration) DeepSeek-V3 38.0 Snow; Reward-SQL 7B PRM
+13.1% BIRD; RL on verified data (ReViSQL) human level; DSGym trajectory fine-tuning +14; explicit state tracking >
more steps (LongDS); retrieved examples + candidate consistency 50 -> 80% (Salesforce).

## Where small models help
Schema linking by cross-encoder (RESDSQL RoBERTa: ROC-AUC 99.4 on Spider; modern: bge-reranker-v2-m3, Qwen3-Reranker);
small SQL specialists (Bird-Fixer 14B ~ o3-mini; 7B PRM +13%). No published numbers found for GLiClass / NLI as SQL
step or result verifiers: an open gap.

## Practitioner consensus
Semantic layer / curated views + glossary; small-model schema pruning; multiple SQL candidates with execution checks and
a selector; explicit clarification; code-based summaries of results, not raw rows in context; logged SQL for audit;
evaluate on your own verified question set.

## Sources
https://arxiv.org/abs/2601.08778 ; https://www.vldb.org/cidrdb/papers/2026/p5-jin.pdf ; https://arxiv.org/abs/2603.20004 ;
https://github.com/uiuc-kang-lab/ReViSQL ; https://bird-bench.github.io/ ; https://arxiv.org/abs/2510.05318 ;
https://arxiv.org/abs/2506.18951 ; https://spider2-sql.github.io/ ; https://github.com/xlang-ai/Spider2 ;
https://haoailab.com/blogs/reforce/ ; https://www.snowflake.com/en/blog/engineering/arctic-text2sql-r1-sql-generation-benchmark/ ;
https://arxiv.org/html/2509.24403v6 ; https://arxiv.org/html/2410.01943v1 ; https://arxiv.org/html/2505.04671 ;
https://arxiv.org/html/2601.16344 ; https://www.together.ai/blog/dsgym ; https://huggingface.co/blog/dabstep ;
https://arxiv.org/html/2506.23719v1 ; https://arxiv.org/abs/2506.06541 ; https://arxiv.org/html/2409.07703v3 ;
https://arxiv.org/html/2407.06423v1 ; https://arxiv.org/html/2410.07331v1 ; https://arxiv.org/pdf/2401.05507 ;
https://arxiv.org/html/2408.09667 ; https://arxiv.org/pdf/2407.01725 ; https://tablebench.github.io/ ;
https://arxiv.org/abs/2605.30434 ; https://arxiv.org/html/2607.01647 ; https://arxiv.org/html/2604.02485v1 ;
https://arxiv.org/pdf/2608.07437 ; https://arxiv.org/pdf/2606.25819 ; https://vllm.ai/blog/2025-10-28-kimi-k2-accuracy ;
https://github.com/MoonshotAI/K2-Vendor-Verifier ; https://www.tinyfish.ai/blog/context-rot ; https://arxiv.org/pdf/2302.05965 ;
https://arxiv.org/pdf/2512.16083 ; https://arxiv.org/html/2501.17174 ; https://docs.getdbt.com/blog/semantic-layer-vs-text-to-sql-2026 ;
https://arxiv.org/pdf/2604.25149 ; https://omni.co/blog/why-text-to-sql-fails ; https://www.getcollate.io/blog/your-text-to-sql-problem-is-not-the-llm ;
https://blog.agami.ai/text-to-sql-accuracy-on-real-company-data/ ; https://kenashe.ai/blog/2026-08-27-correct-sql-answers-can-still-hide-broken-agent-work ;
https://news.ycombinator.com/item?id=38992601 ; https://news.ycombinator.com/item?id=43361333
Gaps: no usable Reddit threads found; no GLiClass/NLI SQL-verifier numbers published.
