# GLiNER and GLiClass as verifier, featurizer and schema-linking helper (research report, 2026-10-03)

Sourcing: HF / arXiv / Reddit blocked for the fetcher; numbers from search extracts of model cards and papers, GitHub
pages and issues (read directly). No Reddit or Discord threads found; community evidence is GitHub issues and PRs.

## Sizes and licences
GLiNER (NAACL 2024): v2.1 small 166M / medium 209M / large 459M / multi 209M, DeBERTa-v3, Apache-2.0 (v0/v1 and
multi_pii-v1 are CC-BY-NC). v2.5 large 435M. Bi-encoder v2.0 (Knowledgator, ModernBERT/Ettin text + sentence-transformer
labels): 60M / 108M / 194M / 530M, "1000+ entity types at near-constant speed" with cached label embeddings. GLiNER2
(Fastino, EMNLP 2025 demo): ~205M, 2,048-token context, NER + classification + extraction + relations, CPU. GLiClass v3
(Apache-2.0, A6000 batch 1): large 439M F1 0.700 (headline 0.719) 25.2 ex/s; base 187M 0.656 51.6 ex/s; modern-large
399M 0.608 43.8 ex/s; modern-base 151M 0.557 54.5 ex/s. v3 is "logic-enhanced" (NLI and logic pre-training + LoRA).

## Hard limits
- Input: GLiNER default max_len 384 words; longer text is silently truncated with only a warning (issue #378: ~70% of
  an article never read). GLiClass DeBERTa: 512 tokens shared by labels and text; ModernBERT versions up to 8k; built-in
  chunking.
- Labels per pass: GLiNER uni-encoder ~50; bi-encoder 1000+ (130x throughput vs uni at 1,024 labels). GLiClass 1 -> 128
  labels: 19.05 -> 17.60 ex/s; an NLI cross-encoder drops 24.55 -> 0.47. GLiNER2 CPU: 130-208 ms for 5-50 labels vs
  1.7-16.9 s for a DeBERTa NLI pipeline.
- Sensitivity: label order changes GLiNER scores (#192); label wording is the main lever; GLiClass saturates (all labels
  0.993-0.9999 on a one-word input, #12); no native "none of the above" (#29).
- Calibration: no published ECE study for either. Signs of miscalibration in #192, #12, #291, #218, #163. GLiNER2-PII
  authors recommend label-specific thresholds or calibration on a small validation set. Treat scores as rankings; fit
  Platt / isotonic on ~100-300 labelled examples per label.
- Speed: GLiNER CPU ~3.7 texts/s (p50 272 ms), ~80x slower than spaCy; 1.5M rows via pandas apply "takes days" (#88)
  -> batch. ONNX slower in one report, quantisation broke accuracy.

## Accuracy
GLiNER-L zero-shot CrossNER+MIT average 60.9 F1 (GLiNER-M ~ UniNER-13B at 140x smaller); 20-dataset: GLiNER-L 47.8,
UniNER-7B 45.7, ChatGPT 36.5. GLiNER2 CrossNER 0.590 vs GPT-4o 0.599. GLiClass large-v3 0.719 vs strongest NLI
cross-encoder 0.682 at 2.3-16x throughput. GLiNER2 vs DeBERTa NLI: SNIPS 0.83 vs 0.77, Banking77 0.70 vs 0.42,
sentiment 0.86-0.87 vs 0.89-0.92; GPT-4o leads by up to ~10. BTZSC (ICLR 2026, 22 datasets): Qwen3-Reranker-8B 0.72,
best LLM 0.67; NLI cross-encoders plateau; GLiClass not confirmed on it (no independent head-to-head).

## Practitioner experience (GitHub anecdotes)
Silent truncation; over-eager extraction (pronouns as PERSON, wide spans); unstable scores; fine-tuning pitfalls
(catastrophic forgetting #163, ModernBERT overfits fast #293, bi-encoder stuck at 0.5 #291, version breakages, GLiClass
broken on transformers v5 #44); domain gaps (research-data NER 0.266 zero-shot -> 0.768 fine-tuned). Tips: chunk at
sentences under ~384 words, descriptive labels, per-label thresholds, batch, bi-encoder with cached labels for big
taxonomies. Production: GLinker, Superlinked SIE, Sim Studio PII, a GLiClass LLM router.

## As a verifier
GLiClass v3 is marketed for NLI, hallucination detection, rule-following verification, reranking (claim/step as label,
evidence as text, one hypothesis per call), but no published hallucination-benchmark number. Comparators: MiniCheck
770M ~GPT-4 on LLM-AggreFact at ~400x lower cost; Bespoke-MiniCheck 77.4% at ~200 ms; HHEM 82.2% accuracy, 8 h -> 10 min
vs an LLM pipeline. Cost ~20-200 ms per (evidence, claim). Design: atomic claims -> GLiClass / MiniCheck -> calibrate on
200-500 own traces -> escalate uncertain to LLM judge.

## Fine-tuning
GLiClass few-shot +0.209 F1 with 8 examples per label; GLiNER2 vendor "3 minutes on 10 examples"; GLiNER-BioMed 10k
LLM-annotated + gold -> +5.96 F1; synthetic data helps but does not replace real labels; few hundred to few thousand
examples, low LR, minutes to <1 h on one GPU; keep general data or LoRA to avoid forgetting.

## Strengths / limits / best uses
| role | strengths | limits | best use |
|---|---|---|---|
| live verifier | one pass scores many hypotheses, 20-200 ms, NLI/logic training, local | uncalibrated, 512-token window shared with labels, no hallucination benchmark numbers, weak multi-hop | cheap first-pass entailment / rule checks on atomic claims; rank plans; calibrate on own traces; escalate low-margin |
| featurizer | labels change without retraining, 25-55 ex/s per stream, per-label scores as features, +0.2 F1 with 8 shots | slow on CPU, needs batching at 1M+ rows, no "none" class, wording matters | multi-label score columns from GLiClass / GLiNER2; keep raw scores, let the TFM learn thresholds |
| schema linking | arbitrary types, bi-encoder 1000+ labels, GLinker linking | no text-to-SQL evaluation published, noisy on numbers/dates, span width 12 | extract value / entity mentions with column-derived labels, link to columns, pass as hints to the SQL LLM |

## Sources
https://github.com/urchade/GLiNER (issues 378, 82, 185, 244, 192, 88, 191, 218, 163, 293, 291, 339; discussion 100) ;
https://github.com/Knowledgator/GLiClass (issues 12, 23, 29) ; https://github.com/Knowledgator/GLinker ;
https://github.com/simstudioai/sim/pull/5495 ; https://github.com/superlinked/sie/pull/414 ; https://github.com/IliasAarab/btzsc ;
https://arxiv.org/abs/2311.08526 ; https://arxiv.org/abs/2508.07662 ; https://arxiv.org/abs/2507.18546 ;
https://arxiv.org/abs/2602.18487 ; https://arxiv.org/abs/2603.11991 ; https://arxiv.org/abs/2605.09973 ;
https://arxiv.org/abs/2504.00676 ; https://arxiv.org/pdf/2605.30582 ; https://huggingface.co/knowledgator/gliclass-large-v3.0 ;
https://huggingface.co/knowledgator/gliner-bi-base-v2.0 ; https://huggingface.co/fastino/gliner2-large-v1 ;
https://docs.knowledgator.com/docs/frameworks/gliclass/pretrained-models/ ; https://docs.bespokelabs.ai/models ;
https://arxiv.org/abs/2512.22416
