import numpy as np
import scipy.sparse as sp

from fer.relbench_hm import hybrid_fill, knn_cf


def test_knn_cf_recovers_group_purchases():
    rng = np.random.default_rng(0)
    n_c, n_i = 200, 50
    group = rng.integers(0, 4, n_c)
    E = np.eye(4)[group] + 0.05 * rng.standard_normal((n_c, 4))
    P = np.zeros((n_c, n_i), np.float32)
    for c in range(n_c):
        P[c, rng.choice(np.arange(group[c] * 10, group[c] * 10 + 10), 3, replace=False)] = 1
    pred = knn_cf(E, sp.csr_matrix(P), np.arange(n_c), k_neighbors=10, K=5)
    assert pred.shape == (n_c, 5)
    assert np.mean([all(group[c] * 10 <= a < group[c] * 10 + 10 for a in pred[c]) for c in range(n_c)]) > 0.95


def test_hybrid_fill_dedups_and_fills():
    out = hybrid_fill(np.array([[1, 2, 2]]), np.array([[2, 3, 4, 5]]), K=4)
    assert out.tolist() == [[1, 2, 3, 4]]
