# Small, fast verifiers as live confidence signals in agent loops (research report, 2026-10-03)

Sourcing: abstracts, model cards and indexed summaries (arxiv / HF not reachable from the fetcher). [recall] = known
result not re-checked; [anecdote] = vendor or blog claim.

## Key findings
- Small verifiers work when trained on in-domain step or claim labels: AgentPRM 3B -> 91% ALFWorld (beats prompted
  GPT-4o); Web-Shepherd 8B +10.9 pts on WebArena-lite for a GPT-4o-mini policy at 1/10 the cost; Weaver's 400M
  cross-encoder keeps 98.7% of a 70B-verifier ensemble's accuracy at 0.03% of the FLOPs; MiniCheck-FT5 (770M) matches
  GPT-4 on grounding at ~400x lower cost.
- Zero-shot, out-of-domain verifiers are weak: ProcessBench F1 Math-Shepherd-7B 31.5, GPT-4o 61.9, o1-mini 87.9;
  best LLM judges ~70% precision on web-trajectory success (AgentRewardBench); NLI zero-shot cross-encoders plateau
  below rerankers and LLMs over 22 datasets (BTZSC; Qwen3-Reranker-8B 0.72 macro-F1).
- Cheap confidence is weak on SQL: BIRD selective prediction AUROC logprob 0.669, self-consistency 0.675, "query
  executes" 0.500, schema relevance 0.553; LLM judges 0.72-0.78; two-provider judge ensemble 0.82 (2026 study).
- Verbalized agent confidence is overconfident: 22% actual vs 77% predicted success; pre-execution assessment
  discriminates better than post-hoc review; deterministic tools (code interpreter) calibrate better than noisy ones.
- Best-supported design: cascade. Cheap verifier settles clear cases, LLM judge takes the uncertain band, thresholds
  calibrated on held-out data (Trust or Escalate: >80% human agreement guaranteed at ~80% coverage; KnowNo conformal
  ask-for-help: 10-24% less human help at target success).

## Encoder judges (size / result)
| model | size | result |
|---|---|---|
| AlignScore | 355M | matches/beats GPT-4-based metrics on 22 datasets |
| MiniCheck-FT5 | 770M | 75.0% balanced acc LLM-AggreFact (GPT-4o ~75.9), ~400x cheaper |
| HHEM-2.1-open | ~110M | 71.8% LLM-AggreFact; near chance on FaithBench (51.4%) |
| LettuceDetect | ModernBERT 8k | 79.2 F1 RAGTruth (+14.8 over Luna); 30-60 ex/s per GPU |
| Granite Guardian | 2-8B | function-call hallucination AUC ~0.79 (0.92 on one set) |
| PromptGuard 2 | 22-86M | AgentDojo attack success 17.6 -> 7.5%, utility -0.7 |
| GLiClass large v3 | encoder | 25.2 ex/s vs 10.6 cross-encoder; 1 -> 128 labels costs 7-20% throughput; zero-shot F1 0.719 (topic/sentiment/intent only) |

## Step-labelled data usable to train or test a step verifier
PRM800K (math); ProcessBench (test); AgentRewardBench (1,302 web trajectories, success/side-effect/repetition);
WebPRM Collection (40K step pairs, best fit); FC-RewardBench (1,500 correct/incorrect function calls; pre-call
validity); Who&When (184 failures, decisive step; methods <15-36%); TRAIL (118 GAIA traces, error taxonomy: natural
GLiClass label set); tau-bench / WebArena / ToolBench (outcome labels; step labels must be derived).

## Pitfalls
Reward hacking under best-of-N (true reward rises then falls with N, sooner for small proxies; Gao et al.);
overconfidence; "query executes" is uninformative; domain shift (Math-Shepherd, HHEM); label-name sensitivity (template
choice can swing zero-shot accuracy from near-perfect to near-chance; average over paraphrases); over-rejection
(LlamaFirewall AlignmentCheck cut utility 47.7 -> 43.1); latency (encoder ms; LLM judge seconds per step); noisy
Monte Carlo step labels; single-run outcome labels unreliable (tau-bench pass^8 ~25% vs pass^1).

## What a GLiClass-style verifier can and cannot do
Can: many-label routing and triage on every step (repeat action, tool error, irreversible action, asks user,
off-task, TRAIL error types); surface guardrails; first stage of a cascade; base for fine-tuning on in-domain step
labels (WebPRM, FC-RewardBench, own traces, SetFit-style); calibrated ask/retry with conformal thresholds.
Cannot: judge zero-shot whether a step is useful or valid; replace a trained claim-support checker (MiniCheck,
LettuceDetect, Granite Guardian); check numeric or aggregate claims (recompute instead); give calibrated scores
out of the box; serve as an optimisation target; see a whole long trajectory.

## Sources
https://arxiv.org/abs/2502.10325 ; https://arxiv.org/html/2511.08325v1 ; https://arxiv.org/abs/2509.11963 ;
https://aclanthology.org/2024.acl-long.510 ; https://arxiv.org/pdf/2305.20050 ; https://arxiv.org/pdf/2410.08146 ;
https://arxiv.org/pdf/2412.06559 ; https://huggingface.co/papers/2504.00891 ; https://arxiv.org/pdf/2506.18203 ;
https://arxiv.org/html/2310.04406v3 ; https://arxiv.org/html/2410.10934v2 ; https://arxiv.org/pdf/2504.08942 ;
https://www.arxiv.org/abs/2505.15277 ; https://aclanthology.org/2024.emnlp-main.499 ; https://bespokelabs.ai/bespoke-minicheck ;
https://aclanthology.org/2023.acl-long.634 ; https://www.vectara.com/blog/hhem-2-1-a-better-hallucination-detection-model/ ;
https://arxiv.org/pdf/2410.13210 ; https://arxiv.org/html/2502.17125v1 ; https://www.ibm.com/granite/docs/models/guardian ;
https://arxiv.org/pdf/2505.03574 ; https://arxiv.org/pdf/2508.07662 ; https://github.com/Knowledgator/GLiClass ;
https://arxiv.org/pdf/2603.11991 ; https://arxiv.org/pdf/2311.07115 ; https://arxiv.org/pdf/2212.10391 ;
https://arxiv.org/pdf/2607.06799 ; https://arxiv.org/html/2411.16742v1 ; https://arxiv.org/pdf/2306.13063 ;
https://www.iclr.cc/virtual/2026/10012609 ; https://arxiv.org/pdf/2601.07264 ; https://arxiv.org/pdf/2307.01928 ;
https://arxiv.org/pdf/2407.18370 ; https://sierra.ai/blog/benchmarking-ai-agents ; https://arxiv.org/pdf/2509.08682 ;
https://arxiv.org/pdf/2505.08638v3 ; https://proceedings.mlr.press/v202/gao23h.html ; https://arxiv.org/pdf/2503.18666
