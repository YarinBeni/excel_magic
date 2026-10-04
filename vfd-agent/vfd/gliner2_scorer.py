"""GLiNER2-family classifiers (fastino/gliner2-*, GLiNER2.5, GLiNER2.5-Decide) behind the same `scores(texts, labels)`
interface as signals.GLiClassScorer: one multi-label task with threshold 0, so every label gets an independent
probability."""
from __future__ import annotations

import numpy as np


class GLiNER2Scorer:
    def __init__(self, model_id: str, device: str = "cuda", batch_size: int = 16):
        try:
            from gliner2 import AutoExtractor
            self.m = AutoExtractor.from_pretrained(model_id)
        except Exception:
            from gliner2 import GLiNER2
            self.m = GLiNER2.from_pretrained(model_id)
        try:
            self.m = self.m.to(device)
        except Exception:
            pass
        self.batch_size = batch_size

    @staticmethod
    def _parse(res, labels) -> np.ndarray:
        out = np.zeros(len(labels), np.float32)
        idx = {lab: j for j, lab in enumerate(labels)}
        items = res.get("verdict", res) if isinstance(res, dict) else res
        if isinstance(items, dict) and "label" in items:
            items = [items]
        for it in items or []:
            if isinstance(it, (tuple, list)) and len(it) == 2:
                lab, p = it
            elif isinstance(it, dict):
                lab, p = it.get("label"), it.get("confidence", it.get("score", 0.0))
            else:
                continue
            if lab in idx:
                out[idx[lab]] = float(p)
        return out

    def scores(self, texts: list[str], labels: list[str]) -> np.ndarray:
        task = {"verdict": {"labels": list(labels), "multi_label": True, "cls_threshold": 0.0}}
        rows = []
        for s in range(0, len(texts), self.batch_size):
            res = self.m.batch_classify_text(texts[s:s + self.batch_size], task, batch_size=self.batch_size,
                                             threshold=0.0, format_results=True, include_confidence=True)
            rows += [self._parse(r, labels) for r in res]
        return np.stack(rows) if rows else np.zeros((0, len(labels)), np.float32)
