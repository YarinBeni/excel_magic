"""End-to-end search harness with harness='none' (no LLM, no model weights): P0 evaluated, held-out scored."""
import json

import pandas as pd

from tabfm_auto.agent.search import run_search
from tabfm_auto.data import load_task


def test_run_search_none_harness(tmp_path):
    t = load_task("iris")
    m = run_search(t, model_spec="logreg", harness="none", baselines=("dummy",), name="t", run_dir=tmp_path / "run",
                   test_size=0.3, seed=0)
    assert m["n_evals"] == 1 and m["p0_cv"] is not None and m["p0_test"] is not None
    assert m["best_test"] == m["p0_test"]
    assert (tmp_path / "run" / "best_pipeline.py").exists()
    assert (tmp_path / "run" / "workspace" / "evals.jsonl").exists()
    assert not (tmp_path / "run" / "workspace" / "test.parquet").exists()  # held-out rows never enter the workspace
    cfg = json.loads((tmp_path / "run" / "config.json").read_text())
    assert cfg["config"]["dataset"] == "iris"


def test_cli_csv_search(tmp_path, monkeypatch):
    from tabfm_auto.cli import main

    t = load_task("wine")
    df = t.X.copy()
    df["label"] = t.y.to_numpy()
    csv = tmp_path / "wine.csv"
    df.to_csv(csv, index=False)
    monkeypatch.chdir(tmp_path)
    rc = main(["search", "--csv", str(csv), "--target", "label", "--harness", "none", "--model", "logreg",
               "--baselines", "dummy", "--description", "wine cultivars"])
    assert rc == 0
    runs = list((tmp_path / "runs").iterdir())
    assert len(runs) == 1 and (runs[0] / "metrics.json").exists()
    assert pd.read_parquet(runs[0] / "workspace" / "train.parquet").shape[0] < len(df)
