"""TabPFN row embeddings with a choosable in-context target and layer (open-tabfm-auto / frozen-embeddings-retrieval).

TEmBed's ``tabpfn`` approach fits TabPFN on the table with an all-zeros label to get label-free row embeddings. On a
synthetic database we found that this target decides what the embedding keeps (segment P@10: random 0.49, k-means
pseudo-labels 0.82). This approach changes only that choice and, optionally, the transformer block read out:

  context_target: zeros (= TEmBed's default) | random | kmeans
  layer: last | <block index 0..11>
"""
import logging
import os

import numpy as np
import pandas as pd
import torch
from benchmark_src.approach_interfaces.base_interface import BaseTabularEmbeddingApproach
from omegaconf import DictConfig

logger = logging.getLogger(__name__)


class TabPFNCtxEmbedder(BaseTabularEmbeddingApproach):
    def __init__(self, cfg: DictConfig):
        super().__init__(cfg)
        self.cfg = cfg
        self.target = str(getattr(cfg.approach, "context_target", "zeros"))
        self.layer = str(getattr(cfg.approach, "layer", "last"))
        self.n_clusters = int(getattr(cfg.approach, "n_clusters", 8))
        self.max_context = int(getattr(cfg.approach, "max_context", 10000))
        self.seed = int(getattr(cfg.approach, "seed", 0))

    def _model(self, n_classes: int):
        from tabpfn import TabPFNClassifier

        ckpt = os.environ.get("TABPFN_CKPT")
        kw = {"model_path": ckpt} if ckpt else {}
        return TabPFNClassifier(device="cuda" if torch.cuda.is_available() else "cpu", n_estimators=1,
                                inference_precision=torch.float32, random_state=self.seed, **kw)

    @staticmethod
    def preprocessing(t: pd.DataFrame) -> pd.DataFrame:
        p = t.copy()
        for c in p.columns:
            if p[c].dtype == "object" or p[c].dtype.name in ("category", "string", "str") or str(p[c].dtype).startswith("string"):
                p[c] = pd.Categorical(p[c].astype(str)).codes
        return p.apply(pd.to_numeric, errors="coerce").fillna(0)

    def _targets(self, X: pd.DataFrame) -> np.ndarray:
        rng = np.random.default_rng(self.seed)
        if self.target == "zeros":
            return np.zeros(len(X), int)
        if self.target == "random":
            return rng.integers(0, 2, len(X))
        if self.target == "kmeans":
            from sklearn.cluster import KMeans
            from sklearn.preprocessing import StandardScaler

            Z = StandardScaler().fit_transform(X.to_numpy(np.float64))
            return KMeans(min(self.n_clusters, max(2, len(X) // 20)), n_init=4, random_state=self.seed).fit_predict(Z)
        raise ValueError(f"unknown context_target {self.target!r}")

    def get_row_embeddings(self, input_table: pd.DataFrame, train_size: int = None, train_labels: np.ndarray = None):
        X = self.preprocessing(input_table)
        y = self._targets(X)
        rng = np.random.default_rng(self.seed)
        ctx = np.sort(rng.choice(len(X), min(len(X), self.max_context), replace=False))
        clf = self._model(len(np.unique(y)))
        clf.fit(X.iloc[ctx], y[ctx])
        captured = []
        handle = None
        if self.layer != "last":
            block = clf.models_[0].blocks[int(self.layer)]
            handle = block.register_forward_hook(lambda m, a, o: captured.append((o[0] if isinstance(o, tuple) else o).detach().float().cpu()))
        out = []
        try:
            for s in range(0, len(X), 2000):
                xc = X.iloc[s:s + 2000]
                captured.clear()
                E = np.asarray(clf.get_embeddings(xc, data_source="test"), np.float32)
                if self.layer == "last":
                    out.append(E.reshape(-1, len(xc), E.shape[-1])[0])
                else:
                    h = captured[-1]
                    out.append(h.reshape(-1, *h.shape[-3:])[0, -len(xc):, -1, :].numpy())
        finally:
            if handle is not None:
                handle.remove()
        E = np.concatenate(out, 0)
        logger.info("tabpfn_ctx target=%s layer=%s -> %s", self.target, self.layer, E.shape)
        return E
