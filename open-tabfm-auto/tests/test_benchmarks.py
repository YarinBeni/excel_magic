import numpy as np
import pandas as pd

from tabfm_auto.benchmarks.elo import elo_ratings
from tabfm_auto.benchmarks.tabarena import compare_to_paper, list_datasets, paper_table9


def test_task_list_and_paper_table_align():
    ds = list_datasets()
    assert len(ds) == 51
    t9 = paper_table9()
    assert len(t9) == 51 and set(t9.dataset) == {d.name for d in ds}
    assert sum(d.n_splits for d in ds) == 816


def test_elo_orders_methods_and_anchors():
    rng = np.random.default_rng(0)
    rows = []
    for t in range(60):
        base = rng.random()
        rows += [{"method": "good", "task": t, "error": base * 0.8}, {"method": "RF", "task": t, "error": base},
                 {"method": "bad", "task": t, "error": base * 1.3}]
    e = elo_ratings(pd.DataFrame(rows), anchor="RF", n_bootstrap=5)
    assert list(e.method) == ["good", "RF", "bad"]
    assert abs(e.set_index("method").loc["RF", "elo"] - 1000) < 1e-6
    assert (e.elo_hi >= e.elo_lo).all()


def test_compare_to_paper():
    res = pd.DataFrame([{"dataset": "credit-g", "pipeline": "P0", "error": 0.20, "status": "ok", "method": "m", "task": "a"},
                        {"dataset": "credit-g", "pipeline": "P*", "error": 0.19, "status": "ok", "method": "m", "task": "a"}])
    c = compare_to_paper(res)
    assert abs(c.loc[0, "our_gain_%"] - 5.0) < 1e-9 and c.loc[0, "tabfm"] == 0.1944
