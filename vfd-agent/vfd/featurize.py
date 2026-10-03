"""G4: turn free-text columns into features a tabular foundation model can use.

  gli        the LLM proposes 8-16 short label names per text column (from the column name and 20 sample values);
             GLiClass scores every row against them; the per-label scores become numeric columns
  tfidf_svd  char/word tf-idf + truncated SVD (what TabPFN's TRANSFORM_TEXT does)
  embed_pca  sentence embedding (bge-small) + PCA
"""
from __future__ import annotations

import json
import re

import numpy as np

LABEL_PROMPT = (
    "A table has a free-text column named '{col}'. Here are sample values:\n{samples}\n\n"
    "{goal}Propose {k} short, distinct category labels (2-5 words each) that describe what these texts are about or "
    "express, such that each text matches one or a few labels. Return only a JSON list of strings.")


def propose_labels(chat, col: str, samples: list[str], k: int = 12, target: str | None = None) -> list[str]:
    goal = (f"The table is used to predict '{target}'. Prefer labels that would help predict it.\n" if target else "")
    shown = "\n".join(f"- {str(s)[:300]}" for s in samples[:20])
    msgs = [{"role": "user", "content": LABEL_PROMPT.format(col=col, samples=shown, k=k, goal=goal)}]
    for _ in range(3):
        text = chat.samples(msgs, n=1, temperature=0.2, max_tokens=600)[0].text
        m = re.search(r"\[.*\]", text, re.S)
        try:
            labels = [str(x).strip() for x in json.loads(m.group(0)) if str(x).strip()] if m else []
        except json.JSONDecodeError:
            labels = []
        labels = list(dict.fromkeys(labels))[:k]
        if len(labels) >= 3:
            return labels
    raise RuntimeError(f"no usable labels for column {col}")


def gli_features(scorer, texts: list[str], labels: list[str], max_chars: int = 1500) -> np.ndarray:
    clean = [str(t)[:max_chars] if isinstance(t, str) and t.strip() else "(empty)" for t in texts]
    return scorer.scores(clean, labels)


def tfidf_svd(train: list[str], test: list[str], dim: int = 32, seed: int = 0) -> tuple[np.ndarray, np.ndarray]:
    from sklearn.decomposition import TruncatedSVD
    from sklearn.feature_extraction.text import TfidfVectorizer

    tr = [str(t) if isinstance(t, str) else "" for t in train]
    te = [str(t) if isinstance(t, str) else "" for t in test]
    v = TfidfVectorizer(min_df=2, max_features=50_000, ngram_range=(1, 2), sublinear_tf=True).fit(tr)
    Xtr, Xte = v.transform(tr), v.transform(te)
    d = max(1, min(dim, Xtr.shape[1] - 1, Xtr.shape[0] - 1))
    svd = TruncatedSVD(d, random_state=seed).fit(Xtr)
    return svd.transform(Xtr).astype(np.float32), svd.transform(Xte).astype(np.float32)


def embed_pca(model, train: list[str], test: list[str], dim: int = 16, seed: int = 0) -> tuple[np.ndarray, np.ndarray]:
    from sklearn.decomposition import PCA

    tr = [str(t)[:2000] if isinstance(t, str) else "" for t in train]
    te = [str(t)[:2000] if isinstance(t, str) else "" for t in test]
    Etr = model.encode(tr, normalize_embeddings=True, batch_size=128)
    Ete = model.encode(te, normalize_embeddings=True, batch_size=128)
    p = PCA(min(dim, Etr.shape[0] - 1), random_state=seed).fit(Etr)
    return p.transform(Etr).astype(np.float32), p.transform(Ete).astype(np.float32)
