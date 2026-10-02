"""Model backends behind one sklearn-like interface (fit / predict / predict_proba on pandas input).

Spec strings: ``"<name>"`` or ``"<name>:k=v,k=v"`` where ``<name>`` is a key of ``tabfm_auto.models.manifest.MODELS``
(or an alias) or one of the classical baselines ``lightgbm``, ``rf``, ``hgb``, ``logreg``, ``dummy``.
Examples: ``tabpfn:n_estimators=8``, ``tabpfn-2.5``, ``tabicl``, ``kumo-tabular-s:n_estimators=4``, ``exaone``.

Switching backbone = changing this string. Weights are resolved by ``tabfm_auto.models.weights``
(``tabfm-models status`` / ``tabfm-models download <name>``).
"""
from __future__ import annotations

import re
from typing import Any

from . import weights as W
from .manifest import MODELS, find

CLASSICAL = ["lightgbm", "rf", "hgb", "logreg", "dummy"]


def parse_model_spec(spec: str) -> tuple[str, dict[str, Any]]:
    name, _, rest = spec.partition(":")
    kw: dict[str, Any] = {}
    for part in filter(None, rest.split(",")):
        k, _, v = part.partition("=")
        try:
            kw[k] = int(v)
        except ValueError:
            try:
                kw[k] = float(v)
            except ValueError:
                kw[k] = {"true": True, "false": False}.get(v.lower(), v)
    return name, kw


def split_model_specs(text: str) -> list[str]:
    """Split a comma-separated list of specs where specs themselves contain commas
    (``"tabpfn:n_estimators=4,device=cuda,hgb"`` -> ``["tabpfn:n_estimators=4,device=cuda", "hgb"]``).
    A token with '=' and no ':' continues the previous spec. ';' is also accepted as a separator."""
    out: list[str] = []
    for tok in re.split(r"[;,]", text):
        tok = tok.strip()
        if not tok:
            continue
        if out and "=" in tok and ":" not in tok:
            out[-1] += "," + tok
        else:
            out.append(tok)
    return out


def list_models() -> list[str]:
    return [m for m in MODELS if MODELS[m].kind == "tabular"] + CLASSICAL


def is_available(spec: str) -> bool:
    name, _ = parse_model_spec(spec)
    if name in CLASSICAL:
        return name != "lightgbm" or W.package_installed("lightgbm")
    try:
        return W.available(name)
    except KeyError:
        return False


# ---------------------------------------------------------------------------------------------
def _tabpfn(card, task_type: str, kw: dict[str, Any]):
    from tabpfn import TabPFNClassifier, TabPFNRegressor

    reg = task_type == "regression"
    if card.name == "tabpfn":
        path = W.WEIGHTS_DIR / ("tabpfn/tabpfn-v2-regressor.ckpt" if reg else "tabpfn/tabpfn-v2-classifier.ckpt")
        if not path.exists():
            raise FileNotFoundError(f"{path} missing: run `tabfm-models download tabpfn`")
        model_path: str = str(path)
    elif card.name in ("tabpfn-3.5", "tabpfn-3.5-fast"):
        model_path = card.hf_files[0]  # single checkpoint for both tasks; tabpfn resolves the version from the name
    else:
        model_path = card.hf_files[1 if reg else 0]
        local = W.WEIGHTS_DIR / "tabpfn" / model_path
        if local.exists():
            model_path = str(local)
    params = {"model_path": model_path, "device": kw.pop("device", "cpu"), "n_estimators": kw.pop("n_estimators", 4),
              "random_state": kw.pop("random_state", 0), "ignore_pretraining_limits": True}
    if isinstance(kw.get("inference_precision"), str) and kw["inference_precision"] not in ("auto", "autocast"):
        import torch

        kw["inference_precision"] = getattr(torch, kw["inference_precision"])  # "float32" -> torch.float32
    params.update(kw)
    return (TabPFNRegressor if reg else TabPFNClassifier)(**params)


def _tabicl(card, task_type: str, kw: dict[str, Any]):
    from tabicl import TabICLClassifier, TabICLRegressor  # type: ignore

    reg = task_type == "regression"
    fname = card.hf_files[1 if reg else 0]
    params: dict[str, Any] = {"device": kw.pop("device", "cpu"), "checkpoint_version": fname,
                              "model_path": str(W.WEIGHTS_DIR / "tabicl" / fname), "allow_auto_download": W.hf_reachable()}
    if "n_estimators" in kw:
        params["n_estimators"] = kw.pop("n_estimators")
    params.update(kw)
    return (TabICLRegressor if reg else TabICLClassifier)(**params)


def _sdm(card, task_type: str, kw: dict[str, Any]):
    from .sdm_wrapper import SdmEstimator

    if card.name.startswith("kumo-tabular"):
        size = {"s": "small", "m": "medium", "l": "large"}[card.name[-1]]
        return SdmEstimator("kumo-tabular", task_type, size=size, **kw)
    return SdmEstimator("tabiclv2", task_type, **kw)


class _NumpyInputs:
    """EXAONE-Tabular wants NumPy arrays; the pipeline runner passes DataFrames. Thin adapter."""

    def __init__(self, est):
        self._est = est

    def fit(self, X, y):
        import numpy as np

        self._est.fit(np.asarray(X, dtype=np.float32), np.asarray(y))
        self.classes_ = getattr(self._est, "classes_", None)
        return self

    def predict(self, X):
        import numpy as np

        return self._est.predict(np.asarray(X, dtype=np.float32))

    def predict_proba(self, X):
        import numpy as np

        return self._est.predict_proba(np.asarray(X, dtype=np.float32))


def _exaone(card, task_type: str, kw: dict[str, Any]):
    from exaonetabular import EXAONETabularClassifier, EXAONETabularRegressor  # type: ignore

    cls = EXAONETabularRegressor if task_type == "regression" else EXAONETabularClassifier
    return _NumpyInputs(cls.from_pretrained(device=kw.pop("device", "cpu"), **kw))


def _sklearn(name: str, task_type: str, kw: dict[str, Any]):
    reg = task_type == "regression"
    if name == "rf":
        from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor

        cls, params = (RandomForestRegressor if reg else RandomForestClassifier), {"n_estimators": 300, "random_state": 0, "n_jobs": -1}
    elif name == "hgb":
        from sklearn.ensemble import HistGradientBoostingClassifier, HistGradientBoostingRegressor

        cls, params = (HistGradientBoostingRegressor if reg else HistGradientBoostingClassifier), {"random_state": 0}
    elif name == "logreg":
        from sklearn.linear_model import LogisticRegression, Ridge
        from sklearn.pipeline import make_pipeline
        from sklearn.preprocessing import StandardScaler

        return make_pipeline(StandardScaler(), Ridge() if reg else LogisticRegression(max_iter=2000))
    elif name == "dummy":
        from sklearn.dummy import DummyClassifier, DummyRegressor

        return DummyRegressor() if reg else DummyClassifier(strategy="prior")
    elif name == "lightgbm":
        import lightgbm as lgb

        cls, params = (lgb.LGBMRegressor if reg else lgb.LGBMClassifier), {"n_estimators": 500, "learning_rate": 0.03,
                                                                            "num_leaves": 31, "verbose": -1,
                                                                            "random_state": 0, "n_jobs": 4}
    else:
        raise KeyError(name)
    params.update(kw)
    return cls(**params)


BUILDERS = {"tabpfn": _tabpfn, "tabicl": _tabicl, "sdm": _sdm, "exaonetabular": _exaone}


def get_model(spec: str, task_type: str):
    """Return an unfitted estimator for ``spec`` (see module docstring)."""
    name, kw = parse_model_spec(spec)
    if name in CLASSICAL:
        return _sklearn(name, task_type, kw)
    card = find(name)
    if card.kind != "tabular":
        raise ValueError(f"{name} is a relational model; use tabfm_auto.relational.embedders")
    st = W.status(card.name)
    if st["state"] != "ready":
        raise RuntimeError(f"model {card.name!r} is not runnable here: {st['state']}. "
                           f"Run `tabfm-models download {card.name}` (needs huggingface.co unless a mirror exists).")
    return BUILDERS[card.package.split(" ")[0]](card, task_type, kw)
