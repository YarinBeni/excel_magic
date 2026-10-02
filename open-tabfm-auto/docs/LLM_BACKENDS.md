# Running the search with (or without) an LLM

The pipeline-search harness is pluggable (`--harness`):

| harness | needs | when |
|---|---|---|
| `heuristic` | nothing beyond the frozen model | no LLM available; a control for "how much of the gain is domain knowledge"; CI |
| `openai` | any OpenAI-compatible endpoint | open-weights coding models on your own GPUs (vLLM / SGLang / Ollama), or OpenAI / OpenRouter |
| `claude-code` | the `claude` CLI logged in or `ANTHROPIC_API_KEY` | the paper's configuration (Claude Code harness) |
| `none` | nothing | score the identity pipeline only |

## No LLM: `--harness heuristic`

Greedy coordinate search over generic operations (sentinel-to-NaN, winsorizing, count encoding of ID-like columns,
log of skewed features, pairwise crosses of the top-k informative features, truncated SVD, column pruning, class-balanced
and multiple random context views, prior correction and temperature in post-processing, log1p target transform,
more ensemble members). This is the paper's "TabFM+" family of operations run as a search; it cannot invent domain
formulas, which is exactly what the LLM adds.

```bash
tabfm-auto search --dataset synth_physics --harness heuristic --budget-evals 20
```

## Open-weights LLM on your cluster: `--harness openai`

Serve a coding model with tool calling, then point the harness at it:

```bash
# one GPU node
vllm serve Qwen/Qwen3-Coder-30B-A3B-Instruct --enable-auto-tool-choice --tool-call-parser hermes --port 8000
# or: vllm serve zai-org/GLM-4.5-Air --enable-auto-tool-choice --tool-call-parser glm45 --port 8000

tabfm-auto search --dataset synth_physics --harness openai \
    --llm Qwen/Qwen3-Coder-30B-A3B-Instruct --llm-base-url http://localhost:8000/v1 --budget-evals 24
```

`OPENAI_BASE_URL` / `OPENAI_API_KEY` are honoured when the flags are omitted. The agent loop
(`tabfm_auto/agent/openai_compat.py`) gives the model five tools: `read_file`, `describe_data`, `write_pipeline`,
`run_eval`, `finish`; every turn is logged to `runs/<id>/agent_stream.jsonl`. Models that support tool calling
well (Qwen3-Coder, GLM-4.5, DeepSeek-V3.x, gpt-oss, Llama-3.3-70B) work out of the box; for a model without a
tool-call parser, serve it with vLLM's `--tool-call-parser` option that matches its template.

## The paper's configuration: `--harness claude-code`

Headless `claude -p` with a restricted tool allow-list (`tabfm_auto/agent/claude_code.py`), `--llm sonnet|opus`.

## Which LLM matters how much?

The paper's ablation (Table 2) attributes +316 Elo to the frozen model's prior and +194 Elo to the LLM search; Opus 5
and Gemini 3.8 Flash reached similar test error with different styles (Opus: fewer, deeper evaluations). Expect an
open 30B-class coder to land between `heuristic` and the frontier models; measure it with
`examples/04_backbone_transfer.py` and the TabArena protocol rather than guessing.

## Free / open coding-agent CLIs: `--harness cli --agent-cmd <preset>`

`tabfm_auto/agent/cli_agent.py` runs any agent CLI headlessly inside the workspace (presets verified 2026-10-02):

| preset | install | provider setup for a vLLM endpoint | notes |
|---|---|---|---|
| `pi` | `npm i -g --ignore-scripts @earendil-works/pi-coding-agent` (Node >= 22.19) | `~/.pi/agent/models.json`: `{"providers":{"vllm":{"baseUrl":"http://HOST:8000/v1","api":"openai-completions","apiKey":"dummy","models":[{"id":"<served name>"}]}}}` | never asks for approval; no turn cap (the eval budget stops it) |
| `qwen-code` | `npm i -g @qwen-code/qwen-code` (Node >= 22) | `OPENAI_BASE_URL`, `OPENAI_API_KEY=x`, `OPENAI_MODEL` (set by the harness from `--llm-base-url/--llm`) | turn / wall-time budgets built in; Qwen OAuth free tier ended 2026-04 |
| `aider` | `pip install aider-chat` | `OPENAI_API_BASE`, `OPENAI_API_KEY=x` | single-turn per call; loop it externally (see `scripts/slurm` J6) |
| `codex` | `npm i -g @openai/codex` | `~/.codex/config.toml` with `model_providers.vllm` (`wire_api = "responses"`, needs vLLM `/v1/responses`) | |
| `openhands` | `uv tool install openhands --python 3.12` | `LLM_MODEL=openai/<name>`, `LLM_BASE_URL`, `LLM_API_KEY=x`; `--override-with-envs` is mandatory | |
| `gemini-cli` | `npm i -g @google/gemini-cli` | `GEMINI_API_KEY` (free key: 250 req/day, Flash) | Gemini models only, no OpenAI-compatible endpoint |

Claude Code cannot be pointed at non-Claude models (Anthropic does not support it through gateways).
vLLM tool-call parsers: Qwen3-Coder `qwen3_coder` (fallback `hermes`), Qwen3 `hermes`, GLM-4.5 `glm45`, gpt-oss `openai`.
