from tabfm_auto.harness.selection import paired_stats, select


def _ev(n, score, folds, status="ok"):
    return {"eval_id": n, "candidate": f"eval_{n:03d}.py", "status": status, "score": score, "folds": folds}


def test_selection_rules():
    p0 = _ev(1, 0.30, [0.31, 0.29, 0.30])
    noise = _ev(2, 0.295, [0.33, 0.25, 0.305])        # better mean, inconsistent folds
    real = _ev(3, 0.27, [0.28, 0.26, 0.27])           # better on every fold by ~0.03
    bad = _ev(4, None, None, status="error")
    evals = [p0, noise, real, bad]
    assert select(evals, "p0") == ["eval_001.py"]
    assert select(evals, "best") == ["eval_003.py"]
    assert select(evals, "gated1") == ["eval_003.py"]
    assert select([p0, noise], "gated1") == ["eval_001.py"]   # the noisy one is rejected -> P0
    assert select([p0, noise], "best") == ["eval_002.py"]
    assert select(evals, "ens3") == ["eval_003.py", "eval_002.py", "eval_001.py"]
    m, se = paired_stats(p0, real)
    assert abs(m - 0.03) < 1e-9 and se < 0.01
