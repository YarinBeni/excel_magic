from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import torch
from sklearn.metrics import roc_auc_score
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.multiclass import OneVsRestClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import make_pipeline
from torch import nn

from .data import RelationalBatch, RelationalTaskGenerator


@dataclass(frozen=True)
class GeneratorConfig:
    context_size: int
    root_cols: int
    child_cols: int
    aux_cols: int
    rows_per_child: int
    rows_per_aux: int
    lag_steps: int
    missing_prob: float
    link_dropout: float
    schema_dropout: float
    task_type: str
    max_classes: int
    seed: int
    binary_quantile_min: float = 0.5
    binary_quantile_max: float = 0.5
    temporal_row_features: bool = True
    entity_history_frac: float = 0.0
    entity_drift_scale: float = 0.3
    root_shortcut_dropout: float = 0.0
    feature_heterogeneity: float = 0.0
    history_label_frac: float = 0.0
    relational_bridge_frac: float = 0.0
    feature_diversity: float = 0.0
    feature_correlation: float = 0.0
    extended_mechanisms: bool = False
    column_permutation: bool = False


def logits_from_batch(model: nn.Module, batch: RelationalBatch) -> torch.Tensor:
    return model(
        batch.root,
        batch.child,
        batch.aux,
        batch.child_mask,
        batch.aux_mask,
        batch.table_mask,
        batch.context_target,
        batch.lag_target,
        batch.time,
        batch.query_mask,
    )


def embeddings_from_batch(model: nn.Module, batch: RelationalBatch) -> torch.Tensor:
    emb = model.encode(
        batch.root,
        batch.child,
        batch.aux,
        batch.child_mask,
        batch.aux_mask,
        batch.table_mask,
        batch.context_target,
        batch.lag_target,
        batch.time,
        batch.query_mask,
    )
    return torch.nan_to_num(emb.float(), nan=0.0, posinf=1e4, neginf=-1e4)


@torch.no_grad()
def collect_predictions(
    model: nn.Module, gen: RelationalTaskGenerator, device: torch.device, batches: int, batch_size: int
) -> tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
    was_training = model.training
    model.eval()
    outputs: list[np.ndarray] = []
    y: list[np.ndarray] = []
    y_class: list[np.ndarray] = []
    task_type: list[np.ndarray] = []
    num_classes: list[np.ndarray] = []
    for _ in range(batches):
        batch = gen.batch(batch_size).to(device)
        out = logits_from_batch(model, batch).float().cpu().numpy()
        out = np.nan_to_num(out, nan=0.0, posinf=1e4, neginf=-1e4)
        outputs.append(out)
        y.append(batch.y.cpu().numpy())
        y_class.append(batch.y_class.cpu().numpy())
        task_type.append(batch.task_type.cpu().numpy())
        num_classes.append(batch.num_classes.cpu().numpy())
    if was_training:
        model.train()
    return (
        np.concatenate(outputs),
        np.concatenate(y),
        np.concatenate(y_class),
        np.concatenate(task_type),
        np.concatenate(num_classes),
    )


@torch.no_grad()
def evaluate_auc(model: nn.Module, gen: RelationalTaskGenerator, device: torch.device, batches: int, batch_size: int) -> float:
    output, y, _, task_type, _ = collect_predictions(model, gen, device, batches, batch_size)
    keep = task_type == 0
    if keep.sum() == 0 or len(np.unique(y[keep])) < 2:
        return float("nan")
    return float(roc_auc_score(y[keep], sigmoid_np(output[keep, 0])))


def evaluate_metrics(model: nn.Module, gen: RelationalTaskGenerator, device: torch.device, batches: int, batch_size: int) -> dict[str, float]:
    output, y, y_class, task_type, num_classes = collect_predictions(model, gen, device, batches, batch_size)
    metrics: dict[str, float] = {}

    binary = task_type == 0
    if binary.sum() > 0 and len(np.unique(y[binary])) > 1:
        metrics["binary_auroc"] = float(roc_auc_score(y[binary], sigmoid_np(output[binary, 0])))
    else:
        metrics["binary_auroc"] = float("nan")

    regression = task_type == 1
    if regression.sum() > 0:
        metrics["regression_mae"] = float(np.abs(output[regression, 1] - y[regression]).mean())
    else:
        metrics["regression_mae"] = float("nan")

    multiclass = task_type == 2
    if multiclass.sum() > 0:
        metrics["multiclass_mrr"] = multiclass_mrr(output[multiclass, 2:], y_class[multiclass], num_classes[multiclass])
    else:
        metrics["multiclass_mrr"] = float("nan")

    score_terms = []
    if np.isfinite(metrics["binary_auroc"]):
        score_terms.append(metrics["binary_auroc"])
    if np.isfinite(metrics["regression_mae"]):
        score_terms.append(1.0 / (1.0 + metrics["regression_mae"]))
    if np.isfinite(metrics["multiclass_mrr"]):
        score_terms.append(metrics["multiclass_mrr"])
    metrics["mean_score"] = float(np.mean(score_terms)) if score_terms else float("nan")
    return metrics


def make_generator(
    config: GeneratorConfig,
    *,
    seed_offset: int = 0,
    mechanism: int | None = None,
    missing_prob: float | None = None,
    link_dropout: float | None = None,
    schema_dropout: float | None = None,
) -> RelationalTaskGenerator:
    return RelationalTaskGenerator(
        context_size=config.context_size,
        root_cols=config.root_cols,
        child_cols=config.child_cols,
        aux_cols=config.aux_cols,
        rows_per_child=config.rows_per_child,
        rows_per_aux=config.rows_per_aux,
        lag_steps=config.lag_steps,
        missing_prob=config.missing_prob if missing_prob is None else missing_prob,
        link_dropout=config.link_dropout if link_dropout is None else link_dropout,
        schema_dropout=config.schema_dropout if schema_dropout is None else schema_dropout,
        mechanism=mechanism,
        task_type=config.task_type,
        max_classes=config.max_classes,
        binary_quantile_min=config.binary_quantile_min,
        binary_quantile_max=config.binary_quantile_max,
        temporal_row_features=config.temporal_row_features,
        entity_history_frac=config.entity_history_frac,
        entity_drift_scale=config.entity_drift_scale,
        root_shortcut_dropout=config.root_shortcut_dropout,
        feature_heterogeneity=config.feature_heterogeneity,
        history_label_frac=config.history_label_frac,
        relational_bridge_frac=config.relational_bridge_frac,
        feature_diversity=config.feature_diversity,
        feature_correlation=config.feature_correlation,
        extended_mechanisms=config.extended_mechanisms,
        column_permutation=getattr(config, 'column_permutation', False),
        seed=config.seed + seed_offset,
    )


def evaluate_mechanisms(
    model: nn.Module,
    config: GeneratorConfig,
    device: torch.device,
    batches: int,
    batch_size: int,
) -> dict[str, float]:
    metrics = {
        f"mechanism_{mechanism}": evaluate_metrics(
            model,
            make_generator(config, seed_offset=10_000 + mechanism * 997, mechanism=mechanism),
            device,
            batches,
            batch_size,
        )["mean_score"]
        for mechanism in range(24)
    }
    metrics["mean"] = finite_mean(metrics.values())
    return metrics


def evaluate_robustness(
    model: nn.Module,
    config: GeneratorConfig,
    device: torch.device,
    batches: int,
    batch_size: int,
) -> dict[str, float]:
    metrics: dict[str, float] = {}
    for missing in (0.0, 0.25, 0.5):
        metrics[f"missing_{missing:.2f}"] = evaluate_metrics(
            model,
            make_generator(config, seed_offset=20_000 + int(missing * 100), missing_prob=missing),
            device,
            batches,
            batch_size,
        )["mean_score"]
    for dropout in (0.0, 0.25, 0.5):
        metrics[f"link_dropout_{dropout:.2f}"] = evaluate_metrics(
            model,
            make_generator(config, seed_offset=30_000 + int(dropout * 100), link_dropout=dropout),
            device,
            batches,
            batch_size,
        )["mean_score"]
    for schema_dropout in (0.0, 0.25, 0.5):
        metrics[f"schema_dropout_{schema_dropout:.2f}"] = evaluate_metrics(
            model,
            make_generator(config, seed_offset=40_000 + int(schema_dropout * 100), schema_dropout=schema_dropout),
            device,
            batches,
            batch_size,
        )["mean_score"]
    metrics["mean"] = finite_mean(metrics.values())
    return metrics


@torch.no_grad()
def evaluate_latent_structure(
    model: nn.Module,
    config: GeneratorConfig,
    device: torch.device,
    batches: int,
    batch_size: int,
) -> dict[str, float]:
    was_training = model.training
    model.eval()
    embeddings: list[np.ndarray] = []
    mechanisms: list[np.ndarray] = []
    table_presence: list[np.ndarray] = []
    gen = make_generator(config, seed_offset=50_000)
    for _ in range(batches):
        batch = gen.batch(batch_size).to(device)
        embeddings.append(embeddings_from_batch(model, batch).cpu().numpy())
        mechanisms.append(batch.mechanism.cpu().numpy())
        table_presence.append(batch.table_mask[:, -1, :].cpu().numpy().astype(np.int64))
    if was_training:
        model.train()

    x = np.concatenate(embeddings)
    y_mech = np.concatenate(mechanisms)
    y_tables = np.concatenate(table_presence)
    metrics: dict[str, float] = {}

    if len(np.unique(y_mech)) > 1 and len(y_mech) >= 12:
        _, counts = np.unique(y_mech, return_counts=True)
        stratify = y_mech if counts.min() >= 2 else None
        x_train, x_test, y_train, y_test = train_test_split(x, y_mech, test_size=0.35, random_state=0, stratify=stratify)
        mechanism_probe = make_pipeline(
            StandardScaler(),
            LogisticRegression(max_iter=500, class_weight="balanced"),
        )
        mechanism_probe.fit(x_train, y_train)
        metrics["mechanism_probe_accuracy"] = float(mechanism_probe.score(x_test, y_test))
    else:
        metrics["mechanism_probe_accuracy"] = float("nan")

    variable_tables = y_tables[:, y_tables.std(axis=0) > 0]
    if len(variable_tables) >= 12 and variable_tables.shape[1] > 0:
        x_train, x_test, y_train, y_test = train_test_split(x, y_tables, test_size=0.35, random_state=1)
        variable = y_train.std(axis=0) > 0
        schema_probe = make_pipeline(
            StandardScaler(),
            OneVsRestClassifier(LogisticRegression(max_iter=500, class_weight="balanced")),
        )
        schema_probe.fit(x_train, y_train[:, variable])
        pred_variable = schema_probe.predict(x_test)
        pred = np.tile(y_train[0], (len(x_test), 1))
        pred[:, variable] = pred_variable
        metrics["schema_probe_exact_match"] = float((pred[:, variable] == y_test[:, variable]).all(axis=1).mean())
        metrics["schema_probe_table_accuracy"] = float((pred[:, variable] == y_test[:, variable]).mean())
    else:
        metrics["schema_probe_exact_match"] = float("nan")
        metrics["schema_probe_table_accuracy"] = float("nan")

    metrics["mean"] = finite_mean(metrics.values())
    return metrics


def finite_mean(values: object) -> float:
    finite = [float(v) for v in values if isinstance(v, float) and np.isfinite(v)]
    return float(np.mean(finite)) if finite else float("nan")


def sigmoid_np(x: np.ndarray) -> np.ndarray:
    return 1.0 / (1.0 + np.exp(-x))


def multiclass_mrr(logits: np.ndarray, y_class: np.ndarray, num_classes: np.ndarray) -> float:
    reciprocal_ranks: list[float] = []
    for row, target, classes in zip(logits, y_class, num_classes):
        active = row[: int(classes)]
        order = np.argsort(-active)
        rank = int(np.where(order == int(target))[0][0]) + 1
        reciprocal_ranks.append(1.0 / rank)
    return float(np.mean(reciprocal_ranks)) if reciprocal_ranks else float("nan")
