import numpy as np

from fer.retrieval_bench import segment_retrieval


def test_segment_retrieval_perfect_embedding():
    seg = np.repeat(np.arange(4), 30)
    E = np.eye(4)[seg] + 0.01 * np.random.default_rng(0).standard_normal((120, 4))
    r = segment_retrieval(E, seg, k=10)
    assert r["P@10"] > 0.99 and r["kNN_acc"] == 1.0


def test_segment_retrieval_random_embedding_is_chance():
    rng = np.random.default_rng(0)
    seg = rng.integers(0, 4, 400)
    r = segment_retrieval(rng.standard_normal((400, 16)), seg, k=10)
    assert abs(r["P@10"] - r["chance_P@10"]) < 0.08
