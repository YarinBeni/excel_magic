from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class PaperTarget:
    suite: str
    task_group: str
    metric: str
    value: float | None
    direction: str
    source: str
    note: str = ""
    expected_tasks: int = 0


PAPER_TARGETS: dict[str, PaperTarget] = {
    "relbenchv1_classification_avg_auroc": PaperTarget(
        suite="RelBenchV1",
        task_group="classification",
        metric="avg_auroc",
        value=79.60,
        direction="higher",
        source="technical_report/KumoRFM-technical-report.pdf Table 3",
        note="KumoRFM-2 row: Avg 79.60, Rank 3.08. Reported as AUROC points, not fraction.",
        expected_tasks=12,
    ),
    "relbenchv1_regression_normalized_mae": PaperTarget(
        suite="RelBenchV1",
        task_group="regression",
        metric="normalized_mae",
        value=0.822,
        direction="lower",
        source="technical_report/KumoRFM-technical-report.pdf Table 7",
        note="KumoRFM-2 row: normalized average MAE mu_n 0.822, Rank 2.67.",
        expected_tasks=9,
    ),
    "relbenchv2_classification_avg_auroc": PaperTarget(
        suite="RelBenchV2",
        task_group="classification",
        metric="avg_auroc",
        value=55.28,
        direction="higher",
        source="technical_report/KumoRFM-technical-report.pdf Table 4 / text around line 899",
        note="Reported as AUROC points, not fraction.",
        expected_tasks=5,
    ),
    "relbenchv2_regression_normalized_mae": PaperTarget(
        suite="RelBenchV2",
        task_group="regression",
        metric="normalized_mae",
        value=0.601,
        direction="lower",
        source="technical_report/KumoRFM-technical-report.pdf Table 8 / text around line 1346",
        expected_tasks=2,
    ),
    "salt_multiclass_mrr_icl": PaperTarget(
        suite="SALT",
        task_group="multiclass",
        metric="avg_mrr",
        value=0.83,
        direction="higher",
        source="technical_report/KumoRFM-technical-report.pdf Table 9",
        expected_tasks=8,
    ),
    "salt_multiclass_mrr_tuned": PaperTarget(
        suite="SALT",
        task_group="multiclass_finetuned",
        metric="avg_mrr",
        value=0.89,
        direction="higher",
        source="technical_report/KumoRFM-technical-report.pdf Table 9",
        expected_tasks=8,
    ),
}


def targets_payload() -> dict[str, Any]:
    return {key: target.__dict__ for key, target in PAPER_TARGETS.items()}


def compare_manifest_to_target(manifest_path: Path | str, target_key: str, metric_path: str | None = None, tolerance: float = 0.05) -> dict[str, Any]:
    manifest_path = Path(manifest_path)
    manifest = json.loads(manifest_path.read_text())
    if target_key not in PAPER_TARGETS:
        raise KeyError(f"Unknown target {target_key}. Available: {sorted(PAPER_TARGETS)}")
    target = PAPER_TARGETS[target_key]
    if metric_path is None:
        metric_path = infer_metric_path(target)
    observed = lookup_path(manifest, metric_path)
    payload: dict[str, Any] = {
        "manifest": str(manifest_path),
        "target_key": target_key,
        "metric_path": metric_path,
        "observed": observed,
        "target": target.__dict__,
        "within_5_percent": None,
        "relative_gap": None,
        "claimable": False,
    }
    if target.value is None:
        payload["reason"] = "target value is not registered; extract the exact paper table value before claiming reproduction"
        return payload
    if observed is None:
        payload["reason"] = "observed metric path is missing"
        return payload
    observed_f = normalize_observed(float(observed), target)
    target_f = float(target.value)
    if target.direction == "higher":
        relative_gap = (target_f - observed_f) / max(abs(target_f), 1e-12)
        within = observed_f >= target_f * (1.0 - tolerance)
    else:
        relative_gap = (observed_f - target_f) / max(abs(target_f), 1e-12)
        within = observed_f <= target_f * (1.0 + tolerance)
    payload["observed_normalized"] = observed_f
    payload["relative_gap"] = float(relative_gap)
    payload["within_5_percent"] = bool(within)
    coverage = infer_coverage(manifest, metric_path)
    payload["coverage_tasks"] = coverage
    payload["expected_tasks"] = target.expected_tasks
    if target.expected_tasks and coverage is not None and coverage < target.expected_tasks:
        payload["claimable"] = False
        payload["reason"] = f"insufficient task coverage: observed {coverage}, expected {target.expected_tasks}"
    elif target.expected_tasks and coverage is None and metric_path.startswith("aggregate."):
        payload["claimable"] = False
        payload["reason"] = "aggregate task coverage is missing"
    else:
        payload["claimable"] = bool(within)
    return payload


def infer_metric_path(target: PaperTarget) -> str:
    if target.metric == "avg_mrr":
        return "model_metrics.multiclass_mrr"
    if target.metric == "avg_auroc":
        return "model_metrics.binary_auroc"
    if target.metric == "normalized_mae":
        return "model_metrics.regression_mae"
    return "validation_score"


def normalize_observed(observed: float, target: PaperTarget) -> float:
    if "AUROC points" in target.note and 0.0 <= observed <= 1.0 and target.value is not None and target.value > 1.0:
        return observed * 100.0
    return observed


def lookup_path(payload: dict[str, Any], path: str) -> Any:
    cur: Any = payload
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur


def infer_coverage(manifest: dict[str, Any], metric_path: str) -> int | None:
    parts = metric_path.split(".")
    if len(parts) >= 3 and parts[0] == "aggregate":
        coverage = lookup_path(manifest, ".".join(parts[:2] + ["num_tasks"]))
        return int(coverage) if coverage is not None else None
    return None
