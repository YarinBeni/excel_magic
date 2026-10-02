import tempfile
from pathlib import Path

from tabfm_auto.data import load_task
from tabfm_auto.harness import evaluate_cv
from tabfm_auto.pipeline import IDENTITY_PIPELINE, load_pipeline, run_pipeline


def _p0():
    p = Path(tempfile.mkdtemp()) / "pipeline.py"
    p.write_text(IDENTITY_PIPELINE)
    return p


def test_identity_pipeline_runs_with_sklearn_backend():
    t = load_task("wine")
    mod = load_pipeline(_p0())
    idx = t.X.sample(frac=1.0, random_state=0).index  # wine is sorted by class
    X, y = t.X.loc[idx], t.y.loc[idx]
    pred, info = run_pipeline(mod, X.iloc[:120], y.iloc[:120], X.iloc[120:], t.task_type, "logreg")
    assert pred.shape == (len(t.X) - 120, 3)
    assert info["n_views"] == 1


def test_cv_reports_error_for_broken_pipeline():
    t = load_task("iris")
    p = _p0()
    p.write_text(IDENTITY_PIPELINE.replace("return X_train, X_test", "return X_train.iloc[:, :0], X_test.iloc[:, :0]"))
    r = evaluate_cv(p, t.X, t.y, t.task_type, "logreg")
    assert r["status"] == "error" and "zero columns" in r["error"]


def test_db_tool_hides_underscore_tables(tmp_path):
    from tabfm_auto.agent import db_tool
    from tabfm_auto.data.synthetic_db import generate

    db = generate(tmp_path / "shop.sqlite", n_customers=50, n_products=20)
    import pytest

    with pytest.raises(SystemExit):
        db_tool.main(["sql", "--db", str(db), "--query", "select * from _hidden_ground_truth"])
