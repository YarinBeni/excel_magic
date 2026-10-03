import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from vfd import bird
from vfd.evaluate import best_of_n, ece, summarize
from vfd.llm import extract_sql
from vfd.schema_link import column_label, evaluate_linking, gold_columns, lexical_scores
from vfd.signals import cheap_signals


@pytest.fixture()
def db(tmp_path: Path) -> Path:
    d = tmp_path / "dev_databases" / "shop"
    d.mkdir(parents=True)
    p = d / "shop.sqlite"
    con = sqlite3.connect(p)
    con.executescript("""
        CREATE TABLE customers (id INTEGER PRIMARY KEY, name TEXT, currency TEXT);
        CREATE TABLE orders (id INTEGER PRIMARY KEY, customer_id INTEGER REFERENCES customers(id), amount REAL);
        INSERT INTO customers VALUES (1,'a','EUR'),(2,'b','CZK'),(3,'c','EUR');
        INSERT INTO orders VALUES (1,1,10.0),(2,1,5.5),(3,2,7.0);
    """)
    con.commit()
    con.close()
    (d / "database_description").mkdir()
    (d / "database_description" / "customers.csv").write_text(
        "original_column_name,column_name,column_description,data_format,value_description\n"
        "currency,currency,payment currency of the customer,text,EUR or CZK\n")
    return p


def test_execute_and_equality(db):
    a = bird.execute(db, "SELECT name FROM customers WHERE currency='EUR'")
    b = bird.execute(db, "SELECT name FROM customers WHERE currency = 'EUR' ORDER BY name DESC")
    c = bird.execute(db, "SELECT name FROM customers")
    assert a.ok and bird.same_result(a, b) and not bird.same_result(a, c)
    bad = bird.execute(db, "SELECT nope FROM customers")
    assert not bad.ok and bad.key == "ERR"
    assert bird.find_db(db.parents[1], "shop") == db


def test_schema_render_uses_descriptions(db):
    s = bird.load_schema(db)
    txt = bird.render_schema(s)
    assert "payment currency of the customer" in txt and "orders.customer_id = customers.id" in txt
    sub = bird.render_schema(s, {("customers", "currency")})
    assert "orders" not in sub and "currency" in sub


def test_extract_sql():
    assert extract_sql("text\n```sql\nSELECT 1;\n```") == "SELECT 1"
    assert extract_sql("Here: SELECT a FROM b;") == "SELECT a FROM b"
    assert extract_sql("```sql\nSELECT 1\n```\nbetter:\n```sql\nSELECT 2\n```") == "SELECT 2"


def test_gold_columns_resolves_aliases(db):
    s = bird.load_schema(db)
    g = gold_columns("SELECT c.name, SUM(o.amount) FROM customers c JOIN orders o ON o.customer_id = c.id "
                     "WHERE currency = 'EUR' GROUP BY c.name", s)
    assert g == {("customers", "name"), ("orders", "amount"), ("orders", "customer_id"), ("customers", "id"),
                 ("customers", "currency")}


def test_linking_lexical(db):
    s = bird.load_schema(db)
    gold = {("customers", "currency")}
    rows, lat = evaluate_linking([(1, "customers who pay in EUR currency", s, gold)], {"lexical": lexical_scores})
    assert rows[0]["recall@5"] == 1.0 and "lexical" in lat
    assert column_label(s.tables["customers"][2]).startswith("customers.currency")


def _fake_candidates():
    rng = np.random.default_rng(0)
    rows = []
    for q in range(40):
        for k in range(5):
            correct = bool(rng.random() < 0.5)
            key = "good" if correct else f"bad{rng.integers(3)}"
            rows.append({"qid": q, "cand": k, "greedy": k == 0, "exec_ok": True, "n_rows": 1, "result_key": key,
                         "mean_logprob": -0.1 if correct else -0.3 + rng.normal(0, 0.05), "gold_ok": True,
                         "correct": correct, "sig_gliclass": 0.7 if correct else 0.4, "sig_judge": float(correct)})
    return pd.DataFrame(rows)


def test_signals_and_summary():
    df = cheap_signals(_fake_candidates())
    assert df.groupby("qid")["sig_self_consistency"].max().between(0, 1).all()
    res = summarize(df)
    assert res["signals"]["sig_judge"]["auroc"] == 1.0
    assert res["signals"]["sig_judge"]["best_of_n"] == res["oracle_accuracy"]
    assert 0 <= res["signals"]["stack_cheap_gli"]["ece_isotonic"] <= 1
    assert [c["judge_share"] for c in res["cascade"]][0] == 0.0
    assert res["cascade"][-1]["auroc"] == 1.0


def test_ece_perfect():
    assert ece([0, 1, 1, 0], [0.0, 1.0, 1.0, 0.0]) == 0.0
    df = pd.DataFrame({"qid": [0, 0, 1, 1], "greedy": [True, False, True, False], "correct": [False, True, True, False],
                       "s": [0.1, 0.9, 0.8, 0.2]})
    assert best_of_n(df, "s") == 1.0


def test_p1_score_and_variants():
    import importlib.util
    spec = importlib.util.spec_from_file_location("p1", Path(__file__).resolve().parents[1] / "experiments" / "p1_text_featurizer.py")
    p1 = importlib.util.module_from_spec(spec); spec.loader.exec_module(p1)  # noqa: E702
    proba = np.array([[0.9, 0.1], [0.2, 0.8], [0.3, 0.7]])
    assert p1.score("roc_auc", ["no", "yes", "yes"], None, proba, ["no", "yes"]) == 1.0
    assert p1.score("acc", [1, 0], np.array([1, 1]), None, [0, 1]) == 0.5
    df = pd.DataFrame({"t": ["a", "b"], "c": ["x", "y"], "n": [1.0, 2.0]})
    feats = pd.DataFrame({"gli_task__t__0": [0.1, 0.9], "embed__t__0": [1.0, 2.0], "tfidf__t__0": [0, 1]})
    m = {"categorical": ["c"], "numerical": ["n"]}
    assert list(p1.variant_frame(df, feats, m, "gli_task_embed").columns) == ["c", "n", "gli_task__t__0", "embed__t__0"]
    assert list(p1.variant_frame(df, feats, m, "base").columns) == ["c", "n"]


def test_featurize_tfidf():
    from vfd.featurize import tfidf_svd
    tr = ["red apple pie", "green apple tart", "blue car fast", "red car slow"] * 5
    a, b = tfidf_svd(tr, ["apple pie", "fast car"], dim=3)
    assert a.shape == (20, 3) and b.shape == (2, 3)
