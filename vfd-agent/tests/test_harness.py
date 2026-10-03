import sqlite3
from pathlib import Path

import numpy as np

from vfd import bird
from vfd.harness import CONFIGS, Tools, answer, link_schema
from vfd.llm import Sample


class FakeChat:
    def __init__(self, sqls):
        self.sqls = list(sqls)
        self.calls = 0

    def samples(self, messages, n=1, temperature=0.0, max_tokens=1024):
        out = []
        for _ in range(n):
            s = self.sqls[self.calls % len(self.sqls)]
            self.calls += 1
            out.append(Sample(f"```sql\n{s}\n```", -0.1, -1.0, 10))
        return out

    def yes_prob(self, messages):
        return (0.9 if "currency = 'EUR'" in messages[-1]["content"] else 0.1), "x"


class FakeGli:
    """Risk labels fire on queries without the currency filter; correctness hypotheses fire on queries with it."""
    def scores(self, texts, labels):
        from vfd.harness import RISK_LABELS
        def one(t, lab):
            has = "currency" in t.split("SQL:")[-1]
            return (0.1 if has else 0.9) if lab in RISK_LABELS else (0.9 if has else 0.1)
        return np.array([[one(t, lab) for lab in labels] for t in texts], np.float32)


class FakeLinker:
    def scores(self, query, labels):
        return np.array([1.0 if "currency" in lab else 0.0 for lab in labels], np.float32)


def make_db(tmp_path: Path) -> Path:
    d = tmp_path / "shop"; d.mkdir()  # noqa: E702
    p = d / "shop.sqlite"
    con = sqlite3.connect(p)
    con.executescript("CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT, currency TEXT, city TEXT);"
                      "INSERT INTO customers VALUES (1,'a','EUR','x'),(2,'b','CZK','y'),(3,'c','EUR','z');")
    con.commit()
    con.close()
    return p


def test_link_keeps_keys(tmp_path):
    p = make_db(tmp_path)
    s = bird.load_schema(p)
    keep = link_schema(s, "currency", Tools(chat=None, linker=FakeLinker()), k=1)
    assert ("customers", "currency") in keep and ("customers", "id") in keep and ("customers", "city") not in keep


def test_configs_run_end_to_end(tmp_path):
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler

    p = make_db(tmp_path)
    s = bird.load_schema(p)
    q = bird.Question(1, "shop", "How many customers pay in EUR?", "", "SELECT count(*) FROM customers WHERE currency = 'EUR'")
    gold = bird.execute(p, q.gold_sql)
    X = np.array([[1, 1, -0.1, 0.8, 0.9], [1, 1, -0.5, 0.2, 0.1], [0, 0, -1, 0, 0.0], [1, 0, -0.3, 0.1, 0.2]])
    sc = StandardScaler().fit(X)
    stack = (sc, LogisticRegression().fit(sc.transform(X), [1, 0, 0, 0]), None)
    for name in ["A", "B", "C", "D", "F", "SC"]:
        chat = FakeChat(["SELECT count(*) FROM customers WHERE currency = 'EUR'", "SELECT count(*) FROM customers"])
        tools = Tools(chat=chat, judge=chat, gliclass=FakeGli(), linker=FakeLinker(), stack=stack)
        r = answer(q, p, s, CONFIGS[name], tools)
        assert r["exec_ok"], name
        assert (r["result_key"] == gold.key), (name, r["sql"])
