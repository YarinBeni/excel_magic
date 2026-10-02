# Design notes from studying TimeCopilot (2026-10-02)

Source: TimeCopilot/timecopilot at v0.0.36 (MIT; "LLMs x Time Series Foundation Models"; pydantic-ai agent that
selects among wrapped forecasters, does not write code; NeurIPS 2025 BERTs workshop paper; #1 CRPS on GIFT-Eval).
Reviewed by a subagent; star counts could not be verified.

## Patterns adopted or planned

| pattern (TimeCopilot) | status here |
|---|---|
| one base class for every wrapped model, unique aliases, fallback model, cache cleanup between models | partial: `tabfm_auto/models/registry.py` returns sklearn-style estimators; a formal `Backbone` protocol + `clean_cache` is on the roadmap |
| pluggable LLM through provider strings (`"openai:gpt-4o"`, Ollama, Bedrock) and an offline stub model for tests | done differently: `--harness openai` takes any OpenAI-compatible endpoint; tests inject a fake client (`tests/test_harnesses.py`) |
| tools return text summaries for the LLM, invalid names raise a retry with the valid options | done in `agent/openai_compat.py` (text summaries); retry-on-invalid is on the roadmap |
| output validator rejecting a result that does not beat the baseline | done implicitly: the harness freezes the best-CV candidate, P0 included, so a search can never end worse than P0 |
| benchmarks as separate projects with replication tests, results page in docs | `tabfm_auto/benchmarks/` + `examples/05_tabarena_protocol.py`; results page `docs/RESULTS.md` |
| model hub page with version constraints | `docs/MODEL_SWITCHING.md` table + `tabfm-models status` |
| docs tested with mktestdocs, mkdocs-material site, per-release changelogs, trusted-publishing release workflow | roadmap (`docs/ROADMAP.md`) |
| README: tagline, badges, 3 capabilities, 30-second quickstart, hello-world with output, provider section, BibTeX | applied |
| pytest markers to keep live / GPU / benchmark tests out of the default run | applied (`pyproject.toml`) |

## Naming verdict
"Tabular Copilot" is a poor fit: "Copilot" is a Microsoft/GitHub mark, TimeCopilot already wraps TabPFN (a "TabularCopilot"
would read as their sibling), and their positioning (chat assistant that picks a model) differs from ours (a coding agent
that evolves a pipeline around a frozen model). Keep `open-tabfm-auto` for the paper lineage; if a product name is wanted
later, prefer a mechanism-stating name (e.g. `tabfm-evolve`).
