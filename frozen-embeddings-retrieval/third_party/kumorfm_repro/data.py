from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import torch


@dataclass(frozen=True)
class RelationalBatch:
    """A batch of in-context relational tasks.

    Shapes:
      root: [batch, samples, root_cols]
      child: [batch, samples, child_rows, child_cols]
      aux: [batch, samples, aux_rows, aux_cols]
      child_mask: [batch, samples, child_rows]
      aux_mask: [batch, samples, aux_rows]
      table_mask: [batch, samples, 3], true for available root/child/aux tables.
      context_target: [batch, samples]
      lag_target: [batch, samples, lag_steps]
      time: [batch, samples, 2], normalized anchor time and local/global flag.
      query_mask: [batch, samples], true for examples whose target is hidden.
      task_type: [batch], 0=binary, 1=regression, 2=multi-class.
      mechanism: [batch], synthetic structural mechanism id.
      num_classes: [batch], active class count for multi-class tasks.
      y: [batch], scalar query target used for binary/regression.
      y_class: [batch], class index for multi-class tasks, otherwise -1.
      target_all: [batch, samples], scalar target for every context/query row.
      y_class_all: [batch, samples], class target for every row, otherwise -1.
    """

    root: torch.Tensor
    child: torch.Tensor
    aux: torch.Tensor
    child_mask: torch.Tensor
    aux_mask: torch.Tensor
    table_mask: torch.Tensor
    context_target: torch.Tensor
    lag_target: torch.Tensor
    time: torch.Tensor
    query_mask: torch.Tensor
    task_type: torch.Tensor
    mechanism: torch.Tensor
    num_classes: torch.Tensor
    y: torch.Tensor
    y_class: torch.Tensor
    target_all: torch.Tensor
    y_class_all: torch.Tensor

    def to(self, device: torch.device | str) -> "RelationalBatch":
        return RelationalBatch(
            root=self.root.to(device),
            child=self.child.to(device),
            aux=self.aux.to(device),
            child_mask=self.child_mask.to(device),
            aux_mask=self.aux_mask.to(device),
            table_mask=self.table_mask.to(device),
            context_target=self.context_target.to(device),
            lag_target=self.lag_target.to(device),
            time=self.time.to(device),
            query_mask=self.query_mask.to(device),
            task_type=self.task_type.to(device),
            mechanism=self.mechanism.to(device),
            num_classes=self.num_classes.to(device),
            y=self.y.to(device),
            y_class=self.y_class.to(device),
            target_all=self.target_all.to(device),
            y_class_all=self.y_class_all.to(device),
        )


class RelationalTaskGenerator:
    """Synthetic relational SCM task generator.

    Each generated task has a root table, a child/event table, and an auxiliary
    table connected through the child rows. Labels are produced by randomly
    sampled structural mechanisms. This intentionally includes mechanisms where
    column-wise pre-aggregation loses information, such as within-row
    co-occurrence tests, plus temporal mechanisms that benefit from lagged
    targets and mixed local/global context examples.
    """

    def __init__(
        self,
        context_size: int = 64,
        root_cols: int = 8,
        child_cols: int = 8,
        aux_cols: int = 6,
        rows_per_child: int = 12,
        rows_per_aux: int = 8,
        lag_steps: int = 4,
        noise_cols: int = 4,
        missing_prob: float = 0.05,
        link_dropout: float = 0.0,
        schema_dropout: float = 0.0,
        mechanism: int | None = None,
        task_type: str = "binary",
        max_classes: int = 8,
        binary_quantile_min: float = 0.5,
        binary_quantile_max: float = 0.5,
        temporal_row_features: bool = True,
        entity_history_frac: float = 0.0,
        entity_drift_scale: float = 0.3,
        root_shortcut_dropout: float = 0.0,
        feature_heterogeneity: float = 0.0,
        history_label_frac: float = 0.0,
        relational_bridge_frac: float = 0.0,
        feature_diversity: float = 0.0,
        feature_correlation: float = 0.0,
        extended_mechanisms: bool = False,
        column_permutation: bool = False,
        seed: int = 0,
    ) -> None:
        self.context_size = context_size
        self.root_cols = root_cols
        self.child_cols = child_cols
        self.aux_cols = aux_cols
        self.rows_per_child = rows_per_child
        self.rows_per_aux = rows_per_aux
        self.lag_steps = lag_steps
        self.noise_cols = noise_cols
        self.missing_prob = missing_prob
        self.link_dropout = link_dropout
        self.schema_dropout = schema_dropout
        self.mechanism = mechanism
        self.task_type = task_type
        self.max_classes = max_classes
        self.binary_quantile_min = binary_quantile_min
        self.binary_quantile_max = binary_quantile_max
        self.temporal_row_features = temporal_row_features
        self.entity_history_frac = float(np.clip(entity_history_frac, 0.0, 1.0))
        self.entity_drift_scale = max(0.0, float(entity_drift_scale))
        self.root_shortcut_dropout = float(np.clip(root_shortcut_dropout, 0.0, 1.0))
        self.feature_heterogeneity = float(np.clip(feature_heterogeneity, 0.0, 1.0))
        self.history_label_frac = float(np.clip(history_label_frac, 0.0, 1.0))
        self.relational_bridge_frac = float(np.clip(relational_bridge_frac, 0.0, 1.0))
        self.feature_diversity = float(np.clip(feature_diversity, 0.0, 1.0))
        self.feature_correlation = float(np.clip(feature_correlation, 0.0, 1.0))
        self.extended_mechanisms = bool(extended_mechanisms)
        self.column_permutation = bool(column_permutation)
        self.rng = np.random.default_rng(seed)
        self._corr_cache: dict[int, np.ndarray] = {}

    @property
    def samples(self) -> int:
        return self.context_size + 1

    def batch(self, batch_size: int) -> RelationalBatch:
        tasks = [self._task() for _ in range(batch_size)]
        root, child, aux, child_mask, aux_mask, table_mask, lag_target, time_feat, target, y_class, task_type, mechanism, num_classes = zip(*tasks)

        root_t = torch.tensor(np.stack(root), dtype=torch.float32)
        child_t = torch.tensor(np.stack(child), dtype=torch.float32)
        aux_t = torch.tensor(np.stack(aux), dtype=torch.float32)
        child_mask_t = torch.tensor(np.stack(child_mask), dtype=torch.bool)
        aux_mask_t = torch.tensor(np.stack(aux_mask), dtype=torch.bool)
        table_mask_t = torch.tensor(np.stack(table_mask), dtype=torch.bool)
        lag_t = torch.tensor(np.stack(lag_target), dtype=torch.float32)
        time_t = torch.tensor(np.stack(time_feat), dtype=torch.float32)
        target_np = np.stack(target).astype(np.float32)
        y_class_np = np.stack(y_class).astype(np.int64)
        task_type_np = np.array(task_type, dtype=np.int64)
        mechanism_np = np.array(mechanism, dtype=np.int64)
        num_classes_np = np.array(num_classes, dtype=np.int64)

        context_target = target_np.copy()
        query_mask = np.zeros_like(context_target, dtype=bool)
        query_mask[:, -1] = True
        context_target[:, -1] = 0.0

        return RelationalBatch(
            root=root_t,
            child=child_t,
            aux=aux_t,
            child_mask=child_mask_t,
            aux_mask=aux_mask_t,
            table_mask=table_mask_t,
            context_target=torch.tensor(context_target, dtype=torch.float32),
            lag_target=lag_t,
            time=time_t,
            query_mask=torch.tensor(query_mask, dtype=torch.bool),
            task_type=torch.tensor(task_type_np, dtype=torch.long),
            mechanism=torch.tensor(mechanism_np, dtype=torch.long),
            num_classes=torch.tensor(num_classes_np, dtype=torch.long),
            y=torch.tensor(target_np[:, -1], dtype=torch.float32),
            y_class=torch.tensor(y_class_np[:, -1], dtype=torch.long),
            target_all=torch.tensor(target_np, dtype=torch.float32),
            y_class_all=torch.tensor(y_class_np, dtype=torch.long),
        )

    def _task(
        self,
    ) -> tuple[
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        np.ndarray,
        int,
        int,
        int,
    ]:
        n = self.samples
        anchor = np.sort(self.rng.uniform(0.05, 1.0, size=n)).astype(np.float32)
        local = self._local_context_flags(n)

        root = self._generate_diverse_features(n, self.root_cols)
        latent = self.rng.normal(size=(n, 3)).astype(np.float32)
        self._inject_entity_history(latent, local, anchor)
        root[:, :3] = 0.65 * root[:, :3] + latent

        child = self._generate_diverse_features(n, self.rows_per_child, self.child_cols)
        child[:, :, :3] += latent[:, None, :]
        aux = self._generate_diverse_features(n, self.rows_per_aux, self.aux_cols)
        aux[:, :, :3] += 0.7 * latent[:, None, :]

        if self.feature_correlation > 0:
            self._apply_feature_correlation(root, protected_tail=0)
            temporal_tail = self._temporal_feature_width(child)
            self._apply_feature_correlation(child, protected_tail=temporal_tail)
            temporal_tail = self._temporal_feature_width(aux)
            self._apply_feature_correlation(aux, protected_tail=temporal_tail)

        if self.noise_cols > 0:
            self._inject_noise_columns(root, protected_tail=0)
            temporal_tail = self._temporal_feature_width(child)
            self._inject_noise_columns(child, protected_tail=temporal_tail)
            temporal_tail = self._temporal_feature_width(aux)
            self._inject_noise_columns(aux, protected_tail=temporal_tail)

        self._inject_temporal_row_features(child, anchor)
        self._inject_temporal_row_features(aux, anchor)
        self._apply_feature_heterogeneity(root, protected_tail=0)
        self._apply_feature_heterogeneity(child, protected_tail=self._temporal_feature_width(child))
        self._apply_feature_heterogeneity(aux, protected_tail=self._temporal_feature_width(aux))

        period = float(self.rng.integers(1, 6))
        periodic = np.sin(anchor * 2.0 * np.pi * period) * 0.1
        root[:, 0] += periodic * 0.4
        child[:, :, 0] += periodic[:, None] * 0.2

        child_mask = np.ones((n, self.rows_per_child), dtype=bool)
        aux_mask = np.ones((n, self.rows_per_aux), dtype=bool)
        if self.link_dropout > 0:
            child_mask &= self.rng.random(size=child_mask.shape) > self.link_dropout
            aux_mask &= self.rng.random(size=aux_mask.shape) > self.link_dropout
            child_mask[:, 0] = True
            aux_mask[:, 0] = True

        table_available = np.ones(3, dtype=bool)
        if self.schema_dropout > 0:
            table_available[1:] = self.rng.random(size=2) > self.schema_dropout
        table_mask = np.broadcast_to(table_available, (n, 3)).copy()
        if not table_available[1]:
            child_mask[:] = False
            child[:] = 0.0
        if not table_available[2]:
            aux_mask[:] = False
            aux[:] = 0.0

        bridge_score: np.ndarray | None = None
        if (
            self.relational_bridge_frac > 0
            and table_available[1]
            and table_available[2]
            and self.rng.random() < self.relational_bridge_frac
        ):
            bridge_score = self._inject_relational_bridge_process(root, child, aux, child_mask, aux_mask, anchor)

        num_mechanisms = 36 if self.extended_mechanisms else 24
        mechanism = self.mechanism if self.mechanism is not None else int(self.rng.integers(0, num_mechanisms))
        if mechanism < 0 or mechanism >= num_mechanisms:
            raise ValueError(f"mechanism must be in [0, {num_mechanisms - 1}], got {mechanism}")
        if mechanism == 0:
            a = child[:, :, 0] + 0.4 * root[:, None, 0] > 0.35
            b = child[:, :, 1] - 0.4 * root[:, None, 1] > 0.35
            score = ((a & b) & child_mask).any(axis=1).astype(np.float32)
        elif mechanism == 1:
            score = (child[:, :, 2] * child[:, :, 3] * child_mask).sum(axis=1)
            score += 0.5 * root[:, 2]
        elif mechanism == 2:
            recent = child[:, -max(2, self.rows_per_child // 3) :, 0]
            recent_mask = child_mask[:, -recent.shape[1] :]
            score = (recent * recent_mask).mean(axis=1) + 0.8 * root[:, 3]
        elif mechanism == 3:
            score = np.maximum(child[:, :, 4] * child_mask, -4).max(axis=1) - root[:, 4]
        elif mechanism == 4:
            score = root[:, 0] * root[:, 1] + (child[:, :, :2].sum(axis=2) * child_mask).mean(axis=1)
        elif mechanism == 5:
            child_signal = self._masked_mean(child[:, :, 0], child_mask)
            aux_signal = self._masked_mean(aux[:, :, 1], aux_mask)
            score = child_signal * aux_signal + 0.3 * root[:, 0]
        elif mechanism == 6:
            early = self._masked_mean(child[:, : self.rows_per_child // 2, 2], child_mask[:, : self.rows_per_child // 2])
            late = self._masked_mean(aux[:, -max(1, self.rows_per_aux // 2) :, 2], aux_mask[:, -max(1, self.rows_per_aux // 2) :])
            score = late - early + 0.7 * anchor + 0.2 * root[:, 2]
        elif mechanism == 7:
            child_gate = child[:, :, 0] + root[:, None, 0] > 0.0
            child_strength = self._masked_mean(child[:, :, 1], child_mask & child_gate)
            aux_gate = aux[:, :, 0] + child_strength[:, None] > 0.0
            score = self._masked_mean(aux[:, :, 2], aux_mask & aux_gate) + 0.2 * root[:, 1]
        elif mechanism == 8:
            child_window = max(1, self.rows_per_child // 3)
            aux_window = max(1, self.rows_per_aux // 3)
            child_recent = self._masked_mean(child[:, -child_window:, 2], child_mask[:, -child_window:])
            aux_recent = self._masked_mean(aux[:, -aux_window:, 3], aux_mask[:, -aux_window:])
            score = child_recent * aux_recent + 0.5 * root[:, 2] + 0.4 * anchor
        elif mechanism == 9:
            # Subgroup discovery: split child rows into two latent groups by col 3 sign,
            # target is difference between group means.
            pos_group = child[:, :, 0] > 0.0
            neg_group = ~pos_group
            pos_mean = self._masked_mean(child[:, :, 1], child_mask & pos_group)
            neg_mean = self._masked_mean(child[:, :, 1], child_mask & neg_group)
            score = (pos_mean - neg_mean) * (1.0 + 0.3 * root[:, 0])
        elif mechanism == 10:
            # Multi-scale temporal: short-term trend vs long-term baseline
            short_n = max(1, self.rows_per_child // 4)
            short = child[:, -short_n:, 2]
            short_mask = child_mask[:, -short_n:]
            long_start = max(0, self.rows_per_child - 3 * short_n)
            long_slice = child[:, long_start:, 2]
            long_mask = child_mask[:, long_start:]
            short_avg = self._masked_mean(short, short_mask)
            long_avg = self._masked_mean(long_slice, long_mask)
            score = (short_avg - long_avg) * 2.0 + 0.5 * root[:, 3] + 0.3 * anchor
        elif mechanism == 11:
            # Nonlinear interaction discovery: specific column pair product
            col_a = child[:, :, 3] * child_mask
            col_b = child[:, :, 4] * child_mask
            interaction = (col_a * col_b).sum(axis=1)
            normalization = (child_mask.sum(axis=1).clip(min=1))
            score = interaction / normalization + 0.2 * root[:, 4] * root[:, 5]
        elif mechanism == 12:
            # Conditional aggregation: mean of child rows where aux row condition holds
            aux_threshold = self._masked_mean(aux[:, :, 1], aux_mask)
            child_condition = child[:, :, 0] > aux_threshold[:, None]
            conditional_mean = self._masked_mean(child[:, :, 2], child_mask & child_condition)
            score = conditional_mean + 0.4 * root[:, 1] + 0.2 * anchor
        elif mechanism == 13:
            # Set comparison: compare first half vs second half of child rows
            half = self.rows_per_child // 2
            first = self._masked_mean(child[:, :half, 5], child_mask[:, :half])
            second = self._masked_mean(child[:, half:, 5], child_mask[:, half:])
            score = (first - second).clip(-3, 3) + np.abs(root[:, 0])
        elif mechanism == 14:
            # Ratio-based: proportion of child rows exceeding a threshold
            threshold = root[:, :3].mean(axis=1)
            exceeds = (child[:, :, 0] > threshold[:, None]) & child_mask
            ratio = exceeds.sum(axis=1).astype(np.float32) / child_mask.sum(axis=1).clip(min=1)
            score = ratio * 3.0 + 0.3 * root[:, 0]
        elif mechanism == 15:
            # Ranking-based: target depends on the rank of root value among child values
            root_val = root[:, 0]
            child_vals = child[:, :, 0]
            n_exceed = (child_vals > root_val[:, None]).sum(axis=1).astype(np.float32)
            total = child_mask.sum(axis=1).clip(min=1)
            rank_fraction = n_exceed / total
            score = rank_fraction * 2.0 + 0.1 * anchor
        elif mechanism == 16:
            # Three-hop metapath: root→child→aux→deep composition
            # First gate child rows by root, then gate aux rows by child summary, then use aux to gate deeper child pattern
            gate1 = child[:, :, 1] + root[:, None, 1] > 0.1
            child_gated_mean = self._masked_mean(child[:, :, 2], child_mask & gate1)
            gate2 = aux[:, :, 1] + child_gated_mean[:, None] > 0.0
            aux_gated_mean = self._masked_mean(aux[:, :, 2], aux_mask & gate2)
            gate3 = child[:, :, 3] + aux_gated_mean[:, None] > 0.0
            score = self._masked_mean(child[:, :, 4], child_mask & gate3) + 0.15 * root[:, 2]
        elif mechanism == 17:
            child_mean = self._masked_mean(child[:, :, 0], child_mask)
            diff_sq = (child[:, :, 0] - child_mean[:, None]) ** 2
            variance = self._masked_mean(diff_sq, child_mask)
            score = np.sqrt(variance.clip(min=1e-6)) + 0.4 * root[:, 0] + 0.2 * anchor
        elif mechanism == 18:
            # Event counting: count child rows where a condition holds, normalized
            condition = (child[:, :, 0] > 0.0) & child_mask
            count = condition.sum(axis=1).astype(np.float32)
            total = child_mask.sum(axis=1).clip(min=1).astype(np.float32)
            score = (count / total) * 2.0 + 0.3 * root[:, 1]
        elif mechanism == 19:
            # Recency: time since last qualifying event (lower = more recent)
            anchor_col = np.tile(anchor[:, None], (1, self.rows_per_child))
            row_time = (np.arange(self.rows_per_child) / max(1, self.rows_per_child - 1)).astype(np.float32)
            row_time = np.tile(row_time[None, :], (n, 1))
            event_time = anchor_col - (1.0 - row_time) * 0.1
            qualifies = (child[:, :, 1] > 0.2) & child_mask
            last_event_time = np.where(qualifies, event_time, -np.inf).max(axis=1)
            has_event = qualifies.any(axis=1)
            score = np.where(has_event, anchor - np.maximum(last_event_time, 0), 0.5) * 2.0 + 0.3 * root[:, 2]
        elif mechanism == 20:
            # Trend: is the mean of recent child rows higher than older rows?
            half = self.rows_per_child // 2
            recent = self._masked_mean(child[:, half:, 3], child_mask[:, half:])
            older = self._masked_mean(child[:, :half, 3], child_mask[:, :half])
            trend = recent - older
            score = trend + 0.5 * root[:, 3] + 0.3 * anchor
        elif mechanism == 21:
            # Peer comparison: entity score vs average of other entities
            entity_signal = self._masked_mean(child[:, :, 4], child_mask)
            peer_mean = entity_signal.mean() * np.ones_like(entity_signal)
            score = (entity_signal - peer_mean) * 2.0 + 0.5 * root[:, 4]
        elif mechanism == 22:
            # Event diversity: number of distinct "types" based on child column range
            col_vals = child[:, :, 5] * child_mask
            col_min = col_vals.min(axis=1)
            col_max = col_vals.max(axis=1)
            diversity = (col_max - col_min).clip(min=0)
            score = diversity + 0.3 * root[:, 5] + 0.2 * anchor
        elif mechanism == 23:
            # Cross-entity interaction via aux table
            child_summary = self._masked_mean(child[:, :, 0], child_mask)
            aux_summary = self._masked_mean(aux[:, :, 0], aux_mask)
            interaction = child_summary * aux_summary
            score = interaction * 2.0 + 0.4 * root[:, 0] + 0.3 * root[:, 1] + 0.2 * anchor
        elif mechanism == 24:
            # Distribution shape comparison: ratio of interquartile range to standard
            # deviation — captures whether a distribution is heavy-tailed or compact.
            child_vals = child[:, :, 0] * child_mask
            q75 = np.quantile(child_vals, 0.75, axis=1)
            q25 = np.quantile(child_vals, 0.25, axis=1)
            iqr = np.maximum(q75 - q25, 1e-6)
            child_mean = self._masked_mean(child[:, :, 0], child_mask)
            abs_dev = np.abs(child[:, :, 0] - child_mean[:, None]) * child_mask
            mad = (abs_dev.sum(axis=1) / child_mask.sum(axis=1).clip(min=1))
            score = (iqr / (mad.clip(min=1e-6))) * 0.8 + 0.3 * root[:, 0] + 0.2 * anchor
        elif mechanism == 25:
            # Correlation detection: target is whether two child features co-vary
            # in the same direction across rows within each sample.
            c0 = self._standardize(child[:, :, 0] * child_mask)
            c1 = self._standardize(child[:, :, 1] * child_mask)
            product = (c0 * c1 * child_mask).sum(axis=1)
            norm = child_mask.sum(axis=1).clip(min=2).astype(np.float32) - 1
            correlation = (product / norm.clip(min=1)).clip(-1, 1)
            score = correlation * 2.0 + 0.5 * root[:, 0] * root[:, 1] + 0.2 * anchor
        elif mechanism == 26:
            # Hierarchical two-level aggregation: child stats feed into aux gating
            # which then feeds back to a different child feature.
            child_level1 = self._masked_mean(child[:, :, 0], child_mask)
            aux_gate = self._masked_mean(aux[:, :, 1], aux_mask) > child_level1 * 0.5
            child_level2 = self._masked_mean(child[:, :, 2], child_mask & aux_gate[:, None])
            score = child_level2 * 1.5 + 0.3 * root[:, 2] + 0.3 * anchor
        elif mechanism == 27:
            # Multiplicative interaction with gating: only use child rows where
            # aux row condition is met, then interact with root features.
            aux_thresh = self._masked_mean(aux[:, :, 0], aux_mask)
            child_cond = child[:, :, 1] > aux_thresh[:, None] * 0.3
            child_filtered = self._masked_mean(child[:, :, 2], child_mask & child_cond)
            interaction = child_filtered * root[:, 3]
            score = interaction * 1.5 + 0.3 * aux_thresh + 0.2 * anchor
        elif mechanism == 28:
            # Changepoint detection: find where child feature trend changes direction.
            half = self.rows_per_child // 2
            trend1 = self._masked_mean(child[:, half:, 0], child_mask[:, half:]) - self._masked_mean(child[:, :half, 0], child_mask[:, :half])
            third = self.rows_per_child // 3
            trend2 = self._masked_mean(child[:, -third:, 0], child_mask[:, -third:]) - self._masked_mean(child[:, :third, 0], child_mask[:, :third])
            changepoint = np.abs(trend1 - trend2)
            score = changepoint * 1.5 + 0.4 * root[:, 4] + 0.15 * anchor
        elif mechanism == 29:
            # Compositional causal chain: root→child→aux→target with three hops.
            child_gate = child[:, :, 0] + root[:, None, 0] > 0.0
            child_mid = self._masked_mean(child[:, :, 1], child_mask & child_gate)
            aux_gate = aux[:, :, 0] + child_mid[:, None] > 0.0
            aux_mid = self._masked_mean(aux[:, :, 1], aux_mask & aux_gate)
            final_gate = child[:, :, 2] + aux_mid[:, None] > 0.0
            score = self._masked_mean(child[:, :, 3], child_mask & final_gate) + 0.2 * root[:, 5]
        elif mechanism == 30:
            # Confounding pattern: common latent affects both child and aux,
            # and the target must disentangle the direct vs confounded paths.
            confounder = root[:, 0] * 0.5 + self.rng.normal(size=n).astype(np.float32) * 0.1
            child_conf = child[:, :, 0] + confounder[:, None] * 0.3
            aux_conf = aux[:, :, 0] + confounder[:, None] * 0.3
            child_direct = self._masked_mean(child_conf, child_mask)
            aux_direct = self._masked_mean(aux_conf, aux_mask)
            score = child_direct * 1.2 - confounder * 0.8 + aux_direct * 0.6 + 0.2 * anchor
        elif mechanism == 31:
            # Moderation: root feature changes the relationship between child and aux.
            moderator = np.tanh(root[:, 1]) * 1.5
            child_signal = self._masked_mean(child[:, :, 0], child_mask)
            aux_signal = self._masked_mean(aux[:, :, 1], aux_mask)
            direct = child_signal + aux_signal
            moderated = child_signal * aux_signal * moderator
            score = direct * 0.6 + moderated * 1.2 + 0.2 * root[:, 2]
        elif mechanism == 32:
            # Outlier sensitivity: target depends on presence of extreme values.
            child_vals = child[:, :, 0] * child_mask
            child_mean = self._masked_mean(child[:, :, 0], child_mask)
            child_std = np.sqrt(self._masked_mean((child[:, :, 0] - child_mean[:, None]) ** 2, child_mask).clip(min=1e-6))
            z_scores = np.abs((child_vals - child_mean[:, None]) / (child_std[:, None] + 1e-6)) * child_mask
            max_z = z_scores.max(axis=1)
            count_extreme = (z_scores > 2.0).sum(axis=1).astype(np.float32) / child_mask.sum(axis=1).clip(min=1)
            score = max_z * 0.5 + count_extreme * 1.5 + 0.3 * root[:, 3]
        elif mechanism == 33:
            # Monotonic trend constraint: target depends on whether child feature
            # consistently increases over time.
            child_time = child[:, :, 0] * child_mask
            ranks = np.argsort(np.argsort(child_time, axis=1), axis=1).astype(np.float32)
            row_idx = np.arange(self.rows_per_child, dtype=np.float32)[None, :]
            monotonicity = ((ranks - row_idx) ** 2 * child_mask).mean(axis=1)
            antimonotonic = ((row_idx - ranks) ** 2 * child_mask).mean(axis=1)
            mono_score = (np.minimum(monotonicity, antimonotonic) / self.rows_per_child).clip(max=1)
            score = (1.0 - mono_score) * 2.5 + 0.3 * root[:, 4] + 0.2 * anchor
        elif mechanism == 34:
            # Subgroup heterogeneity: split by root tertile, compute different
            # aggregations per subgroup, target is the interaction.
            tertile = np.digitize(root[:, 0], np.quantile(root[:, 0], [0.33, 0.67]))
            subgroup_score = np.zeros(n, dtype=np.float32)
            for sg in range(3):
                mask = tertile == sg
                if mask.sum() > 0:
                    weight = [0.3, 0.5, 0.8][sg]
                    subgroup_score[mask] = (
                        self._masked_mean(child[mask][:, :, sg % 3], child_mask[mask]) * weight
                        + self._masked_mean(aux[mask][:, :, (sg + 1) % 3], aux_mask[mask]) * (1 - weight)
                    )
            score = subgroup_score + 0.25 * root[:, 5] + 0.2 * anchor
        elif mechanism == 35:
            # Sequential event pattern: target is whether a specific sequence of
            # child events occurs (A then B within a time window).
            event_a = child[:, :, 0] > 0.3
            event_b = child[:, :, 1] > 0.3
            both = event_a & event_b & child_mask
            a_then_b = np.zeros(n, dtype=np.float32)
            for i in range(1, self.rows_per_child):
                prev_a = event_a[:, i - 1] & child_mask[:, i - 1]
                curr_b = event_b[:, i] & child_mask[:, i]
                a_then_b += (prev_a & curr_b).astype(np.float32)
            total_pairs = (child_mask[:, :-1] & child_mask[:, 1:]).sum(axis=1).clip(min=1).astype(np.float32)
            pattern_score = a_then_b / total_pairs
            score = pattern_score * 3.0 + 0.2 * root[:, 0] + 0.15 * anchor

        score = score.astype(np.float32)
        if bridge_score is not None:
            score = (0.35 * self._standardize(score) + 1.1 * self._standardize(bridge_score)).astype(np.float32)
        score = self._inject_history_label_process(score, root, child, aux, child_mask, aux_mask, anchor, local)
        score = score.astype(np.float32) + 0.15 * self.rng.normal(size=n).astype(np.float32)
        chosen_task_type = self._choose_task_type()
        target, y_class, num_classes = self._target_from_score(score, chosen_task_type)

        lag_target = self._lagged_targets(target, anchor)
        time_feat = np.stack([anchor, local], axis=1).astype(np.float32)
        self._apply_root_shortcut_dropout(root)
        self._apply_missing(root)
        self._apply_missing(child)
        self._apply_missing(aux)
        if not table_available[1]:
            child_mask[:] = False
            child[:] = 0.0
        if not table_available[2]:
            aux_mask[:] = False
            aux[:] = 0.0
        child[~child_mask] = 0.0
        aux[~aux_mask] = 0.0
        if self.column_permutation:
            self._permute_feature_columns(root, child, aux)
        return root, child, aux, child_mask, aux_mask, table_mask, lag_target, time_feat, target, y_class, chosen_task_type, mechanism, num_classes

    def _permute_feature_columns(self, root: np.ndarray, child: np.ndarray, aux: np.ndarray) -> None:
        """Per-task random permutation of genuine feature columns.

        Labels are already generated, so permuting the input columns forces the
        model to infer which columns carry signal from the in-context examples
        rather than memorizing fixed column indices. Temporal row-feature columns
        (appended at the tail of child/aux) keep their position so their learned
        time-since-anchor meaning stays consistent with the eval-time adapter.
        """
        self._permute_one(root, root.shape[-1])
        self._permute_one(child, child.shape[-1] - self._temporal_feature_width(child))
        self._permute_one(aux, aux.shape[-1] - self._temporal_feature_width(aux))

    def _permute_one(self, arr: np.ndarray, n_feat: int) -> None:
        if n_feat <= 1:
            return
        perm = self.rng.permutation(n_feat)
        arr[..., :n_feat] = arr[..., :n_feat][..., perm]

    def _apply_root_shortcut_dropout(self, root: np.ndarray) -> None:
        if self.root_shortcut_dropout <= 0 or self.rng.random() >= self.root_shortcut_dropout:
            return
        root[:] = self.rng.normal(size=root.shape).astype(np.float32)

    def _inject_noise_columns(self, array: np.ndarray, *, protected_tail: int = 0) -> None:
        width = array.shape[-1]
        usable_width = max(0, width - protected_tail)
        if usable_width <= 0:
            return
        noise_cols = min(self.noise_cols, usable_width)
        start = usable_width - noise_cols
        array[..., start:usable_width] = self.rng.normal(size=(*array.shape[:-1], noise_cols)).astype(np.float32)

    def _generate_diverse_features(
        self,
        *dims: int,
    ) -> np.ndarray:
        """Generate features from a diverse set of native distributions.

        When feature_diversity > 0, each column is drawn from a randomly chosen
        distribution family instead of all being standard normal. This reduces
        the gap between synthetic IID-normal features and real relational data
        where columns are binary, skewed, sparse, multi-modal, or heavy-tailed.
        """
        result = self.rng.normal(size=(*dims,)).astype(np.float32)
        if self.feature_diversity <= 0:
            return result

        num_cols = dims[-1]
        for col in range(num_cols):
            if self.rng.random() >= self.feature_diversity:
                continue
            shape = (*dims[:-1],)
            kind = int(self.rng.integers(0, 9))
            if kind == 0:
                # Binary: sign with random threshold
                threshold = float(self.rng.normal(scale=0.5))
                values = self.rng.normal(size=shape).astype(np.float32)
                result[..., col] = np.where(values > threshold, 1.0, -1.0)
            elif kind == 1:
                # Lognormal: right-skewed positive
                sigma = float(self.rng.uniform(0.3, 1.2))
                values = self.rng.lognormal(mean=0.0, sigma=sigma, size=shape).astype(np.float32)
                result[..., col] = self._standardize(values)
            elif kind == 2:
                # Beta-ish via logistic transform: bounded [-1, 1]
                scale = float(self.rng.uniform(0.5, 2.5))
                values = np.tanh(self.rng.normal(scale=scale, size=shape)).astype(np.float32)
                result[..., col] = values
            elif kind == 3:
                # Gamma: positive, skewed
                shape_param = float(self.rng.uniform(0.8, 4.0))
                scale_param = float(self.rng.uniform(0.5, 2.0))
                values = self.rng.gamma(shape=shape_param, scale=scale_param, size=shape).astype(np.float32)
                result[..., col] = self._standardize(values)
            elif kind == 4:
                # Mixture of 2 Gaussians: bimodal
                mu1 = float(self.rng.normal(scale=1.2))
                mu2 = mu1 + float(self.rng.uniform(1.5, 4.0))
                mix = self.rng.random(size=shape) < 0.5
                values = self.rng.normal(scale=0.7, size=shape).astype(np.float32)
                values[mix] += mu2
                values[~mix] += mu1
                result[..., col] = self._standardize(values)
            elif kind == 5:
                # Zero-inflated: sparse with a mass at zero
                zero_prob = float(self.rng.uniform(0.3, 0.7))
                values = self.rng.normal(size=shape).astype(np.float32)
                values[self.rng.random(size=shape) < zero_prob] = 0.0
                result[..., col] = values
            elif kind == 6:
                # Student's t: heavy tails, df > 2 for finite variance
                df = float(self.rng.uniform(2.5, 10.0))
                values = (self.rng.standard_t(df=df, size=shape) / np.sqrt(df / (df - 2))).astype(np.float32)
                result[..., col] = np.clip(values, -8.0, 8.0)
            elif kind == 7:
                # Ordinal: discretized into 3-7 bins
                n_bins = int(self.rng.integers(3, 8))
                values = self.rng.normal(size=shape).astype(np.float32)
                bins = np.quantile(values, np.linspace(0, 1, n_bins + 1)[1:-1])
                discretized = np.digitize(values, bins).astype(np.float32)
                result[..., col] = (discretized - discretized.mean()) / (discretized.std() + 1e-6)
            else:
                # Uniform: flat distribution
                lo = float(self.rng.uniform(-2.0, -0.5))
                hi = float(self.rng.uniform(0.5, 2.0))
                result[..., col] = self.rng.uniform(lo, hi, size=shape).astype(np.float32)
        return result

    def _apply_feature_correlation(
        self,
        array: np.ndarray,
        *,
        protected_tail: int = 0,
    ) -> None:
        """Apply random feature-level correlation via cached Cholesky factors.

        The correlation strength is controlled by ``feature_correlation``.
        Only columns outside the protected tail (temporal features) are
        correlated; protected columns pass through unchanged.
        """
        if self.feature_correlation <= 0:
            return
        usable_width = max(0, array.shape[-1] - protected_tail)
        if usable_width < 2:
            return
        orig_shape = array.shape
        flat = array[..., :usable_width].reshape(-1, usable_width)

        L = self._get_correlation_cholesky(usable_width)
        if L is None:
            return

        correlated = flat @ L.T
        array[..., :usable_width] = correlated.reshape(*orig_shape[:-1], usable_width)

    def _get_correlation_cholesky(self, usable_width: int) -> np.ndarray | None:
        """Return a cached Cholesky factor for the given feature width."""
        if usable_width not in self._corr_cache:
            # Build 8 pre-computed factors and cache one for each width
            self._corr_cache[usable_width] = []
            for _ in range(8):
                random_corr = self.rng.normal(size=(usable_width, usable_width)).astype(np.float32)
                corr = random_corr @ random_corr.T / usable_width
                d = np.sqrt(np.diag(corr).clip(min=1e-6))
                corr = corr / np.outer(d, d)
                np.fill_diagonal(corr, 1.0)
                corr = (
                    (1.0 - self.feature_correlation) * np.eye(usable_width, dtype=np.float32)
                    + self.feature_correlation * corr
                )
                eigvals, eigvecs = np.linalg.eigh(corr)
                eigvals = eigvals.clip(min=1e-6)
                corr = eigvecs @ np.diag(eigvals) @ eigvecs.T
                d = np.sqrt(np.diag(corr).clip(min=1e-6))
                corr = corr / np.outer(d, d)
                try:
                    L = np.linalg.cholesky(corr).astype(np.float32)
                except np.linalg.LinAlgError:
                    continue
                self._corr_cache[usable_width].append(L)
        factors = self._corr_cache[usable_width]
        if not factors:
            return None
        return factors[int(self.rng.integers(0, len(factors)))]

    def _apply_feature_heterogeneity(self, array: np.ndarray, *, protected_tail: int = 0) -> None:
        if self.feature_heterogeneity <= 0:
            return
        usable_width = max(0, array.shape[-1] - protected_tail)
        for col in range(usable_width):
            if self.rng.random() >= self.feature_heterogeneity:
                continue
            values = array[..., col].astype(np.float32, copy=False)
            kind = int(self.rng.integers(0, 5))
            if kind == 0:
                threshold = float(self.rng.normal(scale=0.35))
                transformed = np.where(values > threshold, 1.0, -1.0)
            elif kind == 1:
                bins = np.array([-1.0, -0.35, 0.35, 1.0], dtype=np.float32) + self.rng.normal(scale=0.1)
                transformed = (np.digitize(values, np.sort(bins)).astype(np.float32) - 2.0) / 2.0
            elif kind == 2:
                bias = float(self.rng.normal(scale=0.4))
                transformed = np.exp(np.clip(values + bias, -3.0, 3.0)).astype(np.float32)
                transformed = self._standardize(transformed)
            elif kind == 3:
                keep_prob = float(self.rng.uniform(0.2, 0.65))
                keep = self.rng.random(size=values.shape) < keep_prob
                transformed = np.where(keep, values, 0.0)
            else:
                rate = np.exp(np.clip(0.6 * values + float(self.rng.normal(scale=0.25)), -2.0, 2.0))
                transformed = self.rng.poisson(rate).astype(np.float32)
                transformed = self._standardize(transformed)
            array[..., col] = transformed.astype(np.float32)

    def _inject_history_label_process(
        self,
        score: np.ndarray,
        root: np.ndarray,
        child: np.ndarray,
        aux: np.ndarray,
        child_mask: np.ndarray,
        aux_mask: np.ndarray,
        anchor: np.ndarray,
        local: np.ndarray,
    ) -> np.ndarray:
        if self.history_label_frac <= 0 or self.rng.random() >= self.history_label_frac:
            return score
        local_mask = local.astype(bool)
        if local_mask.sum() <= 1:
            return score

        child_event = self._masked_mean(child[:, :, 0], child_mask)
        aux_event = self._masked_mean(aux[:, :, 0], aux_mask)
        event_signal = self._standardize(child_event + 0.5 * aux_event + 0.25 * root[:, 0])
        base_score = self._standardize(score)

        history_score = np.zeros_like(score, dtype=np.float32)
        state = float(self.rng.normal(scale=0.25))
        seasonal_phase = float(self.rng.uniform(0.0, 2.0 * np.pi))
        sorted_idx = np.argsort(anchor)
        for idx in sorted_idx:
            local_weight = 1.0 if local_mask[idx] else 0.25
            seasonal = 0.25 * np.sin(2.0 * np.pi * anchor[idx] + seasonal_phase)
            innovation = 0.65 * event_signal[idx] + 0.2 * root[idx, 1] + seasonal
            history_score[idx] = local_weight * (0.8 * state + innovation) + (1.0 - local_weight) * innovation
            if local_mask[idx]:
                state = float(np.tanh(history_score[idx]) + self.rng.normal(scale=0.08))

        # Keep the original SCM mechanism visible while forcing some tasks to
        # require same-entity label history rather than static query features.
        return (0.45 * base_score + 1.15 * self._standardize(history_score)).astype(np.float32)

    def _inject_relational_bridge_process(
        self,
        root: np.ndarray,
        child: np.ndarray,
        aux: np.ndarray,
        child_mask: np.ndarray,
        aux_mask: np.ndarray,
        anchor: np.ndarray,
    ) -> np.ndarray:
        """Add a join-like child->aux motif and return its causal score.

        The visible schema has no explicit foreign-key tensor, so we expose a
        learnable soft key in both tables and make the label depend on matched
        child/aux rows. This pressures pre-training toward multi-table
        composition instead of root-only shortcuts while remaining benchmark
        agnostic.
        """

        n = root.shape[0]
        if child.shape[2] < 2 or aux.shape[2] < 2:
            return np.zeros(n, dtype=np.float32)

        num_keys = int(self.rng.integers(3, max(4, min(9, self.rows_per_child + 1))))
        key_centers = np.linspace(-1.0, 1.0, num_keys, dtype=np.float32)
        child_key = self.rng.integers(0, num_keys, size=(n, self.rows_per_child))
        aux_key = self.rng.integers(0, num_keys, size=(n, self.rows_per_aux))

        local_shift = (root[:, 0] > np.median(root[:, 0])).astype(np.int64)
        child_key[:, : max(1, self.rows_per_child // 3)] = (
            child_key[:, : max(1, self.rows_per_child // 3)] + local_shift[:, None]
        ) % num_keys
        aux_key[:, -max(1, self.rows_per_aux // 3) :] = (
            aux_key[:, -max(1, self.rows_per_aux // 3) :] + local_shift[:, None]
        ) % num_keys

        child[:, :, 0] = key_centers[child_key] + self.rng.normal(scale=0.03, size=child_key.shape).astype(np.float32)
        aux[:, :, 0] = key_centers[aux_key] + self.rng.normal(scale=0.03, size=aux_key.shape).astype(np.float32)

        child_value = 0.7 * child[:, :, 1] + 0.25 * root[:, None, 1]
        aux_value = 0.7 * aux[:, :, 1] - 0.2 * root[:, None, 2]
        child[:, :, 1] = child_value.astype(np.float32)
        aux[:, :, 1] = aux_value.astype(np.float32)

        match = child_key[:, :, None] == aux_key[:, None, :]
        pair_mask = child_mask[:, :, None] & aux_mask[:, None, :] & match
        pair_signal = np.tanh(child_value[:, :, None] * aux_value[:, None, :]).astype(np.float32)
        matched_sum = (pair_signal * pair_mask).sum(axis=(1, 2))
        matched_count = pair_mask.sum(axis=(1, 2)).clip(min=1).astype(np.float32)
        matched_mean = matched_sum / matched_count

        child_coverage = (
            (match & aux_mask[:, None, :]).any(axis=2) & child_mask
        ).sum(axis=1).astype(np.float32) / child_mask.sum(axis=1).clip(min=1)
        aux_coverage = (
            (match & child_mask[:, :, None]).any(axis=1) & aux_mask
        ).sum(axis=1).astype(np.float32) / aux_mask.sum(axis=1).clip(min=1)
        temporal = np.sin(2.0 * np.pi * anchor * float(self.rng.integers(1, 5))).astype(np.float32)
        return (1.6 * matched_mean + 0.8 * (child_coverage - aux_coverage) + 0.25 * temporal).astype(np.float32)

    @staticmethod
    def _standardize(values: np.ndarray) -> np.ndarray:
        return ((values - values.mean()) / (values.std() + 1e-6)).astype(np.float32)

    def _local_context_flags(self, n: int) -> np.ndarray:
        local = np.zeros(n, dtype=np.float32)
        if self.entity_history_frac > 0:
            local_count = int(round((n - 1) * self.entity_history_frac))
            local_count = max(1, min(n - 1, local_count))
            local[n - 1 - local_count : n] = 1.0
            return local
        local[-max(1, n // 4) :] = 1.0
        self.rng.shuffle(local[:-1])
        local[-1] = 1.0
        return local

    def _inject_entity_history(self, latent: np.ndarray, local: np.ndarray, anchor: np.ndarray) -> None:
        if self.entity_history_frac <= 0:
            return
        local_mask = local.astype(bool)
        if local_mask.sum() <= 1:
            return
        entity_base = self.rng.normal(size=(3,)).astype(np.float32)
        drift = (self.rng.normal(size=(3,)) * self.entity_drift_scale).astype(np.float32)
        relative_time = (anchor[local_mask] - anchor[-1]).astype(np.float32)[:, None]
        noise = self.rng.normal(scale=0.12, size=(int(local_mask.sum()), 3)).astype(np.float32)
        latent[local_mask] = entity_base[None, :] + relative_time * drift[None, :] + noise

    def _apply_missing(self, array: np.ndarray) -> None:
        if self.missing_prob <= 0:
            return
        missing = self.rng.random(size=array.shape) < self.missing_prob
        array[missing] = 0.0

    def _inject_temporal_row_features(self, table: np.ndarray, anchor: np.ndarray) -> None:
        if not self.temporal_row_features or table.shape[2] == 0:
            return
        rows = table.shape[1]
        width = table.shape[2]
        rank = np.linspace(0.0, 1.0, rows, dtype=np.float32)[None, :]
        max_age = self.rng.uniform(0.05, 0.6, size=(len(anchor), 1)).astype(np.float32)
        jitter = self.rng.uniform(0.0, 0.02, size=(len(anchor), rows)).astype(np.float32)
        age = np.clip((1.0 - rank) * max_age + jitter, 0.0, 1.0)
        row_time = np.clip(anchor[:, None] - age, 0.0, 1.0)
        temporal = [
            row_time,
            age,
            np.log1p(age * 100.0) / np.log1p(100.0),
        ]
        for offset, values in enumerate(reversed(temporal[: min(3, width)]), start=1):
            table[:, :, -offset] = values

    def _temporal_feature_width(self, table: np.ndarray) -> int:
        if not self.temporal_row_features:
            return 0
        return min(3, table.shape[-1])

    def _lagged_targets(self, y: np.ndarray, anchor: np.ndarray) -> np.ndarray:
        lagged = np.zeros((len(y), self.lag_steps), dtype=np.float32)
        for i in range(len(y)):
            for lag in range(self.lag_steps):
                j = i - lag - 1
                lagged[i, lag] = y[j] if j >= 0 and anchor[j] < anchor[i] else 0.0
        return lagged

    def _choose_task_type(self) -> int:
        mapping = {"binary": 0, "regression": 1, "multiclass": 2}
        if self.task_type == "mixed":
            return int(self.rng.integers(0, 3))
        if self.task_type not in mapping:
            raise ValueError(f"task_type must be binary, regression, multiclass, or mixed; got {self.task_type}")
        return mapping[self.task_type]

    def _target_from_score(self, score: np.ndarray, task_type: int) -> tuple[np.ndarray, np.ndarray, int]:
        if task_type == 0:
            q_min = float(np.clip(self.binary_quantile_min, 0.01, 0.99))
            q_max = float(np.clip(self.binary_quantile_max, q_min, 0.99))
            quantile = q_min if q_min == q_max else float(self.rng.uniform(q_min, q_max))
            threshold = np.quantile(score[:-1], quantile)
            target = (score > threshold).astype(np.float32)
            y_class = np.full(len(score), -1, dtype=np.int64)
            return target, y_class, 2
        if task_type == 1:
            mean = score[:-1].mean()
            std = score[:-1].std().clip(min=1e-3)
            target = ((score - mean) / std).astype(np.float32)
            y_class = np.full(len(score), -1, dtype=np.int64)
            return target, y_class, 1

        num_classes = int(self.rng.integers(3, self.max_classes + 1))
        quantiles = np.linspace(0, 1, num_classes + 1)[1:-1]
        bins = np.quantile(score[:-1], quantiles)
        classes = np.digitize(score, bins).astype(np.int64)
        target = (classes / max(1, num_classes - 1)).astype(np.float32)
        return target, classes, num_classes

    @staticmethod
    def _masked_mean(values: np.ndarray, mask: np.ndarray) -> np.ndarray:
        weights = mask.astype(np.float32)
        return (values * weights).sum(axis=1) / weights.sum(axis=1).clip(min=1.0)
