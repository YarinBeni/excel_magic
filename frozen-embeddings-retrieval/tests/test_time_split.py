import numpy as np
import pandas as pd
import pytest

from fer.relbench_layers import time_split


def _train(n_times=40, per=100, seed=0):
    rng = np.random.default_rng(seed)
    t = np.repeat(pd.date_range("2000-01-01", periods=n_times, freq="MS"), per)
    return pd.DataFrame({"t": t, "ent": rng.integers(0, 300, len(t)), "y": rng.integers(0, 2, len(t))})


def test_probe_rows_strictly_after_context():
    tr = _train()
    early, late, info = time_split(tr, "t", n_train=1000)
    assert early["t"].max() < late["t"].min()
    assert len(late) >= 1000 and len(late) - 100 < 1000  # just enough timestamps to hold n_train
    assert len(early) + len(late) == len(tr)
    assert info["ctx_split"] == "time" and info["n_probe_pool"] == len(late)


def test_probe_pool_at_most_about_half():
    tr = _train(n_times=10, per=100)
    early, late, _ = time_split(tr, "t", n_train=4000)  # asks for more than exists: capped at half
    assert len(late) == 500 and len(early) == 500


def test_single_timestamp_refused():
    tr = _train(n_times=1, per=500)
    with pytest.raises(ValueError):
        time_split(tr, "t", n_train=100)
