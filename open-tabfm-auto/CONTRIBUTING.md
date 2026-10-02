# Contributing

```bash
python -m venv .venv && . .venv/bin/activate
pip install -e ".[tabpfn,baselines,dev]"
tabfm-models download tabpfn      # 73 MB from the public GCS mirror
pytest -q && ruff check .
```

- Add a model: one `ModelCard` in `tabfm_auto/models/manifest.py` + one builder in `tabfm_auto/models/registry.py`
  (see `docs/MODEL_SWITCHING.md`).
- Add a dataset: one entry in `tabfm_auto/data/registry.py` returning a `TabularTask`.
- Add an agent harness: implement a function like `tabfm_auto/agent/claude_code.py::run_claude_code` and wire it in
  `tabfm_auto/agent/search.py::run_search`.
- Keep the pipeline contract (`tabfm_auto/pipeline/template.py`) stable: discovered pipelines are meant to be portable.
- Every experiment writes to `runs/` through `RunLogger`; never report a number that is not in a `metrics.json`.
