"""Load and validate a candidate ``pipeline.py``."""
from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path

REQUIRED = {"preprocess": 3, "engineer": 3, "sample": 4, "postprocess": 3}
MAX_COLUMNS = 500


class PipelineError(RuntimeError):
    pass


def load_pipeline(path: str | Path) -> types.ModuleType:
    path = Path(path)
    if not path.exists():
        raise PipelineError(f"pipeline file not found: {path}")
    name = f"candidate_pipeline_{abs(hash(str(path.resolve())))}"
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    try:
        spec.loader.exec_module(mod)
    except Exception as e:
        raise PipelineError(f"import of {path.name} failed: {type(e).__name__}: {e}") from e
    validate_pipeline(mod)
    return mod


def validate_pipeline(mod: types.ModuleType) -> None:
    import inspect

    for fn, nargs in REQUIRED.items():
        f = getattr(mod, fn, None)
        if f is None or not callable(f):
            raise PipelineError(f"pipeline.py must define {fn}()")
        n = len([p for p in inspect.signature(f).parameters.values()
                 if p.default is inspect.Parameter.empty and p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD)])
        if n > nargs:
            raise PipelineError(f"{fn}() takes {nargs} positional args, found {n} required")
    kw = getattr(mod, "MODEL_KWARGS", {})
    if not isinstance(kw, dict):
        raise PipelineError("MODEL_KWARGS must be a dict")
