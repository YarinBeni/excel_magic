"""Dataset registry.

Sources that work offline (bundled with scikit-learn) are the default so experiments can run
inside a network-restricted sandbox. OpenML / TabArena loaders are included but need
``api.openml.org`` to be reachable.
"""
from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

import numpy as np
import pandas as pd

TaskType = str  # "binary" | "multiclass" | "regression"


@dataclass
class TabularTask:
    name: str
    X: pd.DataFrame
    y: pd.Series
    task_type: TaskType
    metadata: dict[str, Any] = field(default_factory=dict)

    @property
    def n_classes(self) -> int | None:
        return None if self.task_type == "regression" else int(self.y.nunique())

    def summary(self) -> dict[str, Any]:
        return {
            "name": self.name, "task_type": self.task_type, "n_rows": int(len(self.X)),
            "n_cols": int(self.X.shape[1]), "n_classes": self.n_classes,
            "target": self.y.name,
        }


def infer_task_type(y: pd.Series) -> TaskType:
    if pd.api.types.is_float_dtype(y) and y.nunique() > 20:
        return "regression"
    return "binary" if y.nunique() == 2 else "multiclass"


# ---------------------------------------------------------------------------------------------
# bundled sklearn datasets (no network)
# ---------------------------------------------------------------------------------------------
def _sk(name: str) -> TabularTask:
    from sklearn import datasets as skd

    loaders: dict[str, tuple[Callable, TaskType, str]] = {
        "iris": (skd.load_iris, "multiclass", "Iris flower species from sepal/petal measurements (cm)."),
        "wine": (skd.load_wine, "multiclass", "Wine cultivar from 13 chemical analysis measurements."),
        "breast_cancer": (skd.load_breast_cancer, "binary",
                          "Wisconsin breast cancer: malignant(0)/benign(1) from cell-nucleus image statistics."),
        "digits": (skd.load_digits, "multiclass", "8x8 handwritten digit pixels (0-16 gray levels)."),
        "diabetes": (skd.load_diabetes, "regression",
                     "Diabetes progression one year after baseline from 10 standardized clinical features."),
    }
    loader, ttype, desc = loaders[name]
    b = loader(as_frame=True)
    X: pd.DataFrame = b.data.copy()
    y: pd.Series = b.target.copy()
    y.name = "target"
    meta = {"description": desc, "source": f"sklearn.datasets.load_{name}",
            "columns": {c: "" for c in X.columns}}
    if hasattr(b, "target_names") and ttype != "regression":
        meta["target_names"] = [str(t) for t in b.target_names]
    if hasattr(b, "DESCR"):
        meta["long_description"] = str(b.DESCR)[:4000]
    return TabularTask(name, X, y, ttype, meta)


# ---------------------------------------------------------------------------------------------
# synthetic datasets with planted domain formulas (to check that an agent can rediscover them)
# ---------------------------------------------------------------------------------------------
def _synthetic_physics(n: int = 500, seed: int = 0, noise: float = 0.08) -> TabularTask:
    """Regression where the target is a nonlinear function of physically-named columns.

    target = log(frequency * chord / velocity) * thickness + 0.3*sin(angle) + noise.
    A frozen TFM sees raw columns; an agent that engineers the Strouhal-like ratio should win.
    """
    rng = np.random.default_rng(seed)
    f = rng.uniform(200, 20000, n)
    u = rng.uniform(30, 80, n)
    c = rng.uniform(0.02, 0.3, n)
    a = rng.uniform(0, 22, n)
    d = rng.uniform(0.0004, 0.06, n)
    st = f * c / u
    y = np.log(st) * (1 + 10 * d) + 0.3 * np.sin(np.radians(a)) + noise * rng.standard_normal(n)
    X = pd.DataFrame({"frequency_hz": f, "free_stream_velocity_m_s": u, "chord_length_m": c,
                      "attack_angle_deg": a, "suction_side_thickness_m": d})
    meta = {"description": "Synthetic aeroacoustic-style regression. The response depends on the "
                           "Strouhal-like ratio frequency*chord/velocity, scaled by thickness, plus an "
                           "angle term.", "source": "synthetic", "columns": {c_: "" for c_ in X.columns}}
    return TabularTask("synth_physics", X, pd.Series(y, name="target"), "regression", meta)


def _synthetic_entities(n: int = 1500, seed: int = 0) -> TabularTask:
    """Binary task on high-cardinality IDs where co-occurrence/degree features carry the signal.

    Mirrors the Amazon_employee_access pattern from the paper: raw IDs are useless to a TFM, but
    counts of (role, resource) pairs and per-manager degrees are predictive.
    """
    rng = np.random.default_rng(seed)
    n_roles, n_res, n_mgr = 40, 120, 60
    role = rng.integers(0, n_roles, n)
    mgr = rng.integers(0, n_mgr, n)
    # resources are drawn with a per-role popularity so pair counts become informative
    pop = rng.dirichlet(np.ones(n_res) * 0.3, size=n_roles)
    res = np.array([rng.choice(n_res, p=pop[r]) for r in role])
    pair_count = pd.Series(list(zip(role, res))).map(pd.Series(list(zip(role, res))).value_counts())
    mgr_deg = pd.Series(mgr).map(pd.Series(mgr).value_counts())
    logit = 2.2 * np.log1p(pair_count.values) - 0.8 * np.log1p(mgr_deg.values) - 1.6
    y = (rng.random(n) < 1 / (1 + np.exp(-logit))).astype(int)
    X = pd.DataFrame({"role_id": role, "manager_id": mgr, "resource_id": res,
                      "request_id": rng.permutation(n)})
    meta = {"description": "Synthetic access-request approval. Columns are categorical identifiers. "
                           "Approval depends on how often a (role, resource) pair occurs and on the "
                           "manager's number of requests.", "source": "synthetic",
            "columns": {c: "identifier" for c in X.columns}}
    return TabularTask("synth_entities", X, pd.Series(y, name="approved"), "binary", meta)


# ---------------------------------------------------------------------------------------------
# OpenML / TabArena (needs network)
# ---------------------------------------------------------------------------------------------
def _openml(name: str) -> TabularTask:
    from sklearn.datasets import fetch_openml

    b = fetch_openml(name=name, as_frame=True, parser="auto")
    X, y = b.data, b.target
    ttype = infer_task_type(y)
    if ttype != "regression":
        y = pd.Series(pd.factorize(y)[0], name=y.name, index=y.index)
    return TabularTask(name, X, y, ttype, {"description": str(b.DESCR)[:4000], "source": "openml",
                                            "columns": {c: "" for c in X.columns}})


REGISTRY: dict[str, Callable[[], TabularTask]] = {
    "iris": lambda: _sk("iris"),
    "wine": lambda: _sk("wine"),
    "breast_cancer": lambda: _sk("breast_cancer"),
    "digits": lambda: _sk("digits"),
    "diabetes": lambda: _sk("diabetes"),
    "synth_physics": _synthetic_physics,
    "synth_entities": _synthetic_entities,
}


def list_tasks() -> list[str]:
    return sorted(REGISTRY)


def load_task(name: str) -> TabularTask:
    if name in REGISTRY:
        return REGISTRY[name]()
    if name.startswith("openml:"):
        return _openml(name.split(":", 1)[1])
    if name.startswith("db:"):
        from .synthetic_db import load_db_task

        return load_db_task(name.split(":", 1)[1])
    raise KeyError(f"unknown task {name!r}; known: {list_tasks()} or openml:<name> or db:<task>")
