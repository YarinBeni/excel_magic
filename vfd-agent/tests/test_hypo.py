from types import SimpleNamespace

import duckdb
import numpy as np
import pandas as pd

from vfd.hypo import HypoLoop, hint


def _env(seed=0):
    rng = np.random.default_rng(seed)
    users = np.arange(400)
    ev = pd.DataFrame({"user": rng.choice(users, 20000),
                       "ts": pd.Timestamp("2020-01-01") + pd.to_timedelta(rng.integers(0, 400, 20000), "D")})
    def rows(t0, n):
        r = pd.DataFrame({"user": rng.choice(users, n), "t": pd.Timestamp(t0)})
        m = r.merge(ev, on="user")
        m = m[(m.ts < m.t) & (m.ts >= m.t - pd.Timedelta(30, "D"))]
        cnt = m.groupby(["user", "t"]).size()
        r["c"] = r.set_index(["user", "t"]).index.map(cnt).fillna(0).to_numpy()
        r["y"] = (r["c"] + rng.normal(0, 1.0, n) > r["c"].median()).astype(int)
        return r
    prb, val, test = rows("2020-08-01", 800), rows("2020-10-01", 800), rows("2020-12-01", 800)
    con = duckdb.connect()
    con.register("events", ev)
    noise = lambda n: rng.uniform(0.3, 0.7, n)
    fm = {"prb": prb[["user", "t", "y"]].assign(fm=noise(800)), "val": val[["user", "t", "y"]].assign(fm=noise(800)),
          "test": test[["user", "t"]].assign(fm=noise(800))}
    from sklearn.metrics import roc_auc_score
    task = SimpleNamespace(evaluate=lambda s: {"roc_auc": roc_auc_score(test["y"], s)})
    env = SimpleNamespace(ent="user", tcol="t", tgt="y", con=con, _fm=fm, task=task, train=prb[["user", "t", "y"]],
                          db=SimpleNamespace(table_dict={"events": SimpleNamespace(df=ev, time_col="ts")}),
                          tables={"events": ["user", "ts"]})
    env._q = lambda sql, timeout=120: con.execute(sql).fetchdf()
    return env


GOOD = ("SELECT r.user, r.t, count(e.ts) AS n30 FROM rows r LEFT JOIN events e ON e.user = r.user "
        "AND e.ts < r.t AND e.ts >= r.t - INTERVAL 30 DAY GROUP BY r.user, r.t")
LEAK = ("SELECT r.user, r.t, count(e.ts) AS n_all FROM rows r LEFT JOIN events e ON e.user = r.user "
        "GROUP BY r.user, r.t")


def test_hypothesis_gate_and_leak():
    env = _env()
    h = HypoLoop(env, chat=None, n_boot=100)
    assert h.test_hypothesis("leaky", LEAK).startswith("REJECTED (leakage)")
    out = h.test_hypothesis("recent activity", GOOD)
    assert out.startswith("ACCEPTED"), out
    res = h.finalize(0, 0.0)
    assert res["auroc"] > res["auroc_fm"] + 0.1 and res["accepted"] == ["recent activity"]
    assert "date_diff" in hint("Catalog Error: Scalar Function with name julianday does not exist!")
