# Roadmap

- [ ] Formal `Backbone` protocol (alias, fit/predict/predict_proba, clean_cache) and `fallback_model` in the runner
- [ ] Provider strings for the LLM harness (`openai:...`, `anthropic:...`, `ollama:...`) on top of the current OpenAI-compatible loop
- [ ] Retry-with-valid-options on malformed tool calls; stricter pipeline validation messages for the agent
- [ ] mkdocs-material site, docs tested with mktestdocs, per-release changelogs
- [ ] Release workflow: tag -> build -> wheel/sdist import test -> PyPI trusted publishing
- [ ] TabArena results page once the cluster run exists (`docs/experiments/tabarena.md`)
- [ ] Kumo Tabular / TabICLv2 / EXAONE backbones exercised in CI with cached weights
- [ ] Codex / other coding-agent harnesses (the paper's second and third configurations)
