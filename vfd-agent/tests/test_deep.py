"""DEEP tool on synthetic data with known answers, using fast sklearn backbones (no GPU)."""
import numpy as np
import pandas as pd
import pytest

pytest.importorskip("sklearn")
from vfd.deep import DeepTool  # noqa: E402

try:
    from vfd.deep import _registry
    _registry()("hgb", "binary")
    HAVE = True
except Exception:
    HAVE = False
pytestmark = pytest.mark.skipif(not HAVE, reason="open-tabfm-auto registry not importable")


def data(n=1500, seed=0):
    rng = np.random.default_rng(seed)
    df = pd.DataFrame({"id": np.arange(n), "t": pd.date_range("2024-01-01", periods=n, freq="h"),
                       "price": rng.normal(10, 2, n), "region": rng.choice(["north", "south"], n),
                       "noise": rng.normal(0, 1, n)})
    logit = 1.5 * (df.price - 10) + 1.0 * (df.region == "north")
    df["churn"] = (rng.random(n) < 1 / (1 + np.exp(-logit))).astype(int)
    return df


T = DeepTool(model="hgb", fast="hgb", check="hgb", max_context=2000)


def test_predict_and_drivers():
    df = data()
    r = T.predict(df, "churn", time_col="t", id_col="id", query=df.tail(50))
    assert r["metric"] == "auroc" and r["value"] > 0.75 and "check_value" in r and len(r["query_top"]) == 20
    d = T.drivers(df, "churn", id_col="id", time_col="t")
    assert d["drivers"][0]["column"] == "price"
    assert {x["column"] for x in d["drivers"][:2]} == {"price", "region"}


def test_what_if_monotone():
    w = T.what_if(data(), "churn", "price", values=[6.0, 10.0, 14.0], id_col="id", time_col="t")
    m = [c["mean"] for c in w["curve"]]
    assert m[0] < m[1] < m[2]


def test_anomalies_find_planted():
    df = data()
    df.loc[[5, 50, 500], "price"] = 30.0
    df.loc[[5, 50, 500], "churn"] = 0  # very high price but stayed: surprising
    a = T.anomalies(df, "churn", id_col="id", time_col="t", top=30)
    assert len({5, 50, 500} & {x["id"] for x in a["top"]}) >= 2


def test_drift_detects_shift_and_column():
    old, new = data(seed=1), data(seed=2)
    new["price"] = new["price"] + 3
    r = DeepTool(fast="hgb").drift(old, new, exclude=["id", "t", "churn"])
    assert r["changed"] and r["columns"][0]["column"] == "price"
    same = DeepTool(fast="hgb").drift(data(seed=3), data(seed=4), exclude=["id", "t", "churn"])
    assert same["auroc"] < 0.6


def test_hypothesis_and_similar():
    h = T.hypothesis(data(), "churn", ["price"], id_col="id", time_col="t", repeats=4)
    assert h["supported"] and h["delta_mean"] > 0.05
    h0 = T.hypothesis(data(), "churn", ["noise"], id_col="id", time_col="t", repeats=4)
    assert not h0["supported"]
    E = np.array([[1, 0], [0.9, 0.1], [0, 1], [0.1, 0.9]])
    s = DeepTool.similar(E, ["a", "b", "c", "d"], ["a"], k=2)
    assert s["neighbours"][0]["id"] == "b"
