"""Catalogue of every backbone the lab can run, with where its weights live.

TabArena Elo values are copied from arXiv 2609.37989 Table 4 (51-dataset pool, RandomForest = 1000) where
available, otherwise read approximately from the TabArena frontier chart of 2026-09-29 (marked "~").
They are here so a run's metrics can be put next to the model's expected strength.
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ModelCard:
    name: str                      # spec name used in --model / EMBEDDERS
    kind: str                      # "tabular" | "relational"
    package: str                   # python package that implements it
    hf_repo: str | None            # Hugging Face repo id (None if not on HF)
    hf_files: tuple[str, ...] = () # files needed from the repo (empty = whole snapshot / package decides)
    hf_revision: str | None = None
    mirror_urls: tuple[str, ...] = ()  # non-HF download URLs (GCS / GitHub releases); same order as hf_files
    local_files: tuple[str, ...] = ()  # expected paths under weights/ once downloaded
    params_m: float | None = None
    license: str = ""
    elo_tabarena: str = ""
    cpu_ok: str = ""
    notes: str = ""
    aliases: tuple[str, ...] = field(default_factory=tuple)


_GCS = "https://storage.googleapis.com/tabpfn-v2-model-files/05152025"

MODELS: dict[str, ModelCard] = {m.name: m for m in [
    # ----------------------------------------------------------------------------- tabular
    ModelCard("tabpfn", "tabular", "tabpfn", "Prior-Labs/TabPFN-v2-clf",
              ("tabpfn-v2-classifier.ckpt", "tabpfn-v2-regressor.ckpt"),
              mirror_urls=(f"{_GCS}/tabpfn-v2-classifier.ckpt", f"{_GCS}/tabpfn-v2-regressor.ckpt"),
              local_files=("tabpfn/tabpfn-v2-classifier.ckpt", "tabpfn/tabpfn-v2-regressor.ckpt"),
              params_m=11, license="Prior Labs License (Apache-2.0 + attribution)",
              elo_tabarena="not in pool; below RealTabPFN-2.5 (1522)", cpu_ok="yes (<=10k rows)",
              notes="Only model with a non-HF mirror; the regressor file is on HF as Prior-Labs/TabPFN-v2-reg.",
              aliases=("tabpfn-2",)),
    ModelCard("tabpfn-2.5", "tabular", "tabpfn", "Prior-Labs/tabpfn_2_5",
              ("tabpfn-v2.5-classifier-v2.5_default.ckpt", "tabpfn-v2.5-regressor-v2.5_default.ckpt"),
              params_m=10.7, license="TabPFN-2.5 license (non-commercial)", elo_tabarena="1522 (RealTabPFN-2.5)",
              cpu_ok="yes"),
    ModelCard("tabpfn-2.6", "tabular", "tabpfn", "Prior-Labs/tabpfn_2_6",
              ("tabpfn-v2.6-classifier-v2.6_default.ckpt", "tabpfn-v2.6-regressor-v2.6_default.ckpt"),
              license="TabPFN-2.6 license (non-commercial)", elo_tabarena="1605", cpu_ok="yes"),
    ModelCard("tabpfn-3", "tabular", "tabpfn", "Prior-Labs/tabpfn_3",
              ("tabpfn-v3-classifier-v3_default.ckpt", "tabpfn-v3-regressor-v3_default.ckpt"),
              params_m=53, license="non-commercial", elo_tabarena="1660", cpu_ok="slow (default limit 5000 rows)"),
    ModelCard("tabpfn-3.5", "tabular", "tabpfn", "Prior-Labs/tabpfn_3_5", ("tabpfn-v3.5-20260909.safetensors",),
              license="academic + evaluation", elo_tabarena="~1855", cpu_ok="slow"),
    ModelCard("tabpfn-3.5-fast", "tabular", "tabpfn", "Prior-Labs/tabpfn_3_5", ("tabpfn-v3.5-fast-20260909.safetensors",),
              license="academic + evaluation", elo_tabarena="~1770", cpu_ok="probably"),
    ModelCard("tabicl", "tabular", "tabicl", "jingang/TabICL",
              ("tabicl-classifier-v2-20260212.ckpt", "tabicl-regressor-v2-20260212.ckpt"),
              local_files=("tabicl/tabicl-classifier-v2-20260212.ckpt", "tabicl/tabicl-regressor-v2-20260212.ckpt"),
              params_m=28, license="BSD-3", elo_tabarena="1590", cpu_ok="yes (GPU recommended)",
              notes="In the paper's transfer set (+143 Elo with the discovered pipelines).", aliases=("tabiclv2",)),
    ModelCard("tabiclv2-sdm", "tabular", "sdm", "nvidia/TabICLv2", (),
              params_m=28, license="BSD-3", elo_tabarena="1590", cpu_ok="yes",
              notes="TabICLv2 through NVIDIA structured-data-models; repo id resolved by the package."),
    ModelCard("kumo-tabular-s", "tabular", "sdm", "nvidia/Kumo-Tabular", ("small/classifier.pt", "small/regressor.pt"),
              hf_revision="v1.0.0", params_m=28, license="OpenMDW-1.1 (commercial OK)", elo_tabarena="~1790",
              cpu_ok="expected (TabICLv2-size)", notes="Best licence/size/Elo trade-off for CPU."),
    ModelCard("kumo-tabular-m", "tabular", "sdm", "nvidia/Kumo-Tabular", ("medium/classifier.pt", "medium/regressor.pt"),
              hf_revision="v1.0.0", license="OpenMDW-1.1", elo_tabarena="~1910", cpu_ok="slow"),
    ModelCard("kumo-tabular-l", "tabular", "sdm", "nvidia/Kumo-Tabular", ("large/classifier.pt", "large/regressor.pt"),
              hf_revision="v1.0.0", params_m=215, license="OpenMDW-1.1", elo_tabarena="~1960 (#1)", cpu_ok="impractical"),
    ModelCard("exaone", "tabular", "exaonetabular", "LG-AI-Research/EXAONE-Tabular", (),
              params_m=21, license="code BSD-3; weights non-commercial", elo_tabarena="1765", cpu_ok="slow",
              notes="pip install 'exaonetabular @ git+https://github.com/LGAI-Research/EXAONE-Tabular.git'"),
    # ----------------------------------------------------------------------------- relational
    ModelCard("openrfm", "relational", "kumorfm_repro (vendored)", None,
              mirror_urls=("https://github.com/T-Lab/OpenRFM/releases/download/v0.1.0/openrfm-pretrain-ctx96.pt",),
              local_files=("openrfm/openrfm-pretrain-ctx96.pt",), params_m=19.5, license="MIT",
              elo_tabarena="RelBench 12-task AUROC ~69 (KumoRFM-2: 79.6)", cpu_ok="yes"),
    ModelCard("kumo-relational", "relational", "sdm", "nvidia/Kumo-Relational", ("classifier.pt", "regressor.pt"),
              params_m=30, license="OpenMDW-1.1", elo_tabarena="KumoRFM-2 adapted", cpu_ok="expected",
              notes="4 readout tokens x 128 ch = 512-d graph embedding, captured by a forward hook on the ICL head."),
    ModelCard("rt-j", "relational", "rt",
              "stanford-star/rt-j", (), params_m=22, license="see repo", elo_tabarena="RelBench zero-shot 93% of supervised",
              cpu_ok="eager flex_attention", notes="pip install git+https://github.com/stanford-star/relational-transformer; needs RT's preprocessed tensor format + a sentence-transformers text model (also HF)."),
]}


def find(name: str) -> ModelCard:
    base = name.split(":")[0]
    if base in MODELS:
        return MODELS[base]
    for m in MODELS.values():
        if base in m.aliases:
            return m
    raise KeyError(f"unknown model {name!r}; known: {sorted(MODELS)}")
