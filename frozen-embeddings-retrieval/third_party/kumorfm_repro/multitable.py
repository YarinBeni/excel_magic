"""N-table (database-native) relational foundation model.

Definitive finding from the 3-table effort: the residual RelBench gap is bounded by
the fixed root->child->aux chain, which cannot represent star schemas (e.g. rel-stack:
users <- {badges, votes, posts, comments}). Three child-selection heuristics all capped
at ~68-69 because one child + one aux cannot hold the multiple sibling tables different
tasks need simultaneously.

This module generalizes the model to an arbitrary set of FK->root event tables (N typed
tables), the paper's "database-native" design. It reuses the existing TableEncoder and
GraphCrossSampleEncoder (which already operate over a variable table dimension), so the
change is additive and leaves the 3-table path untouched.
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import torch
from torch import nn

from .model import TableEncoder, GraphCrossSampleEncoder


@dataclass(frozen=True)
class MultiTableBatch:
    """A batch of in-context tasks over a root entity plus N event tables.

    root:        [B, S, Cr]
    events:      [B, S, N, R, Ce]   (N event tables, R rows each, padded)
    event_mask:  [B, S, N, R]       true for real rows
    table_mask:  [B, S, N+1]        true for available tables (root + N events)
    context_target/lag_target/time/query_mask/task_type/num_classes/y/y_class as usual.
    """

    root: torch.Tensor
    events: torch.Tensor
    event_mask: torch.Tensor
    table_mask: torch.Tensor
    context_target: torch.Tensor
    lag_target: torch.Tensor
    time: torch.Tensor
    query_mask: torch.Tensor
    task_type: torch.Tensor
    num_classes: torch.Tensor
    y: torch.Tensor
    y_class: torch.Tensor

    def to(self, device) -> "MultiTableBatch":
        f = lambda t: t.to(device)
        return MultiTableBatch(
            f(self.root), f(self.events), f(self.event_mask), f(self.table_mask),
            f(self.context_target), f(self.lag_target), f(self.time), f(self.query_mask),
            f(self.task_type), f(self.num_classes), f(self.y), f(self.y_class),
        )


class MultiTableRFM(nn.Module):
    """Database-native RFM over a root table and a variable number of event tables."""

    def __init__(
        self,
        root_cols: int,
        event_cols: int,
        num_event_tables: int,
        lag_steps: int = 4,
        max_classes: int = 8,
        d_model: int = 256,
        heads: int = 8,
        table_layers: int = 2,
        graph_layers: int = 3,
        dropout: float = 0.1,
        typed_task_conditioning: bool = False,
    ) -> None:
        super().__init__()
        task_cols = 2 + lag_steps + 2
        self.num_event_tables = num_event_tables
        self.root_encoder = TableEncoder(root_cols + task_cols, d_model, heads, table_layers, dropout)
        # One shared encoder across all event tables (column-position embeddings make it
        # table-agnostic); table identity is supplied by the graph encoder's type embedding.
        self.event_encoder = TableEncoder(event_cols + task_cols, d_model, heads, table_layers, dropout)
        self.graph_encoder = GraphCrossSampleEncoder(num_event_tables + 1, d_model, heads, graph_layers, dropout)
        # Typed task conditioning: project the appended task-role columns (visible target,
        # query/visibility marker, lagged-target history, time/local-context) into a learned
        # conditioning vector added to every table embedding BEFORE graph attention. This is
        # the strongest representation change found in the 3-table effort (+~8 AUROC points);
        # the N-table path previously lacked it entirely. Default off for old-checkpoint compat.
        self.typed_task_conditioning = typed_task_conditioning
        if typed_task_conditioning:
            self.task_cond = nn.Sequential(
                nn.Linear(task_cols, d_model), nn.GELU(), nn.Linear(d_model, d_model)
            )
            self.role_emb = nn.Parameter(torch.zeros(num_event_tables + 1, d_model))
        self.max_classes = max_classes
        self.readout = nn.Sequential(
            nn.LayerNorm(d_model),
            nn.Linear(d_model, d_model),
            nn.GELU(),
            nn.Linear(d_model, 2 + max_classes),
        )

    def forward(self, batch: MultiTableBatch) -> torch.Tensor:
        return self.readout(self.encode(batch))

    def encode(self, batch: MultiTableBatch) -> torch.Tensor:
        b, s, n, r, ce = batch.events.shape
        target_feat = batch.context_target.unsqueeze(-1)
        visible_feat = (~batch.query_mask).float().unsqueeze(-1)
        task_feat = torch.cat([target_feat, visible_feat, batch.lag_target, batch.time], dim=-1)

        root_in = torch.cat([batch.root, task_feat], dim=-1).unsqueeze(2)
        root_emb = self.root_encoder(root_in)

        table_embs = [root_emb]
        for t in range(n):
            ev = batch.events[:, :, t]                       # [B,S,R,Ce]
            ev_extra = task_feat[:, :, None, :].expand(-1, -1, r, -1)
            ev_in = torch.cat([ev, ev_extra], dim=-1)
            ev_emb = self.event_encoder(ev_in, row_mask=batch.event_mask[:, :, t])
            table_embs.append(ev_emb)

        if self.typed_task_conditioning:
            cond = self.task_cond(task_feat)                 # [B,S,d_model]
            table_embs = [emb + cond + self.role_emb[i] for i, emb in enumerate(table_embs)]

        sample_emb = self.graph_encoder(table_embs, table_mask=batch.table_mask)
        query_index = batch.query_mask.float().argmax(dim=1)
        bi = torch.arange(b, device=batch.root.device)
        return sample_emb[bi, query_index]


class MultiTableTaskGenerator:
    """Synthetic generator with a root entity and N event tables sharing a latent.

    Labels depend on a cross-table interaction among a RANDOM SUBSET of the event
    tables, so the model must learn to (a) identify which tables carry signal from
    context and (b) aggregate across multiple tables — exactly the capability a
    fixed 3-table chain lacks on star schemas.
    """

    def __init__(
        self,
        context_size: int = 24,
        root_cols: int = 8,
        event_cols: int = 8,
        num_event_tables: int = 5,
        rows_per_event: int = 10,
        lag_steps: int = 4,
        noise_cols: int = 3,
        missing_prob: float = 0.1,
        link_dropout: float = 0.15,
        table_dropout: float = 0.3,
        task_type: str = "mixed",
        max_classes: int = 8,
        binary_quantile_min: float = 0.5,
        binary_quantile_max: float = 0.5,
        root_shortcut_dropout: float = 0.0,
        history_label_frac: float = 0.0,
        seed: int = 0,
    ) -> None:
        self.context_size = context_size
        self.root_cols = root_cols
        self.event_cols = event_cols
        self.num_event_tables = num_event_tables
        self.rows_per_event = rows_per_event
        self.lag_steps = lag_steps
        self.noise_cols = noise_cols
        self.missing_prob = missing_prob
        self.link_dropout = link_dropout
        self.table_dropout = table_dropout
        self.task_type = task_type
        self.max_classes = max_classes
        self.binary_quantile_min = binary_quantile_min
        self.binary_quantile_max = binary_quantile_max
        self.root_shortcut_dropout = root_shortcut_dropout
        self.history_label_frac = history_label_frac
        self.rng = np.random.default_rng(seed)

    @property
    def samples(self) -> int:
        return self.context_size + 1

    def batch(self, batch_size: int) -> MultiTableBatch:
        tasks = [self._task() for _ in range(batch_size)]
        (root, events, ev_mask, tab_mask, lag, time_feat, target, y_class, ttype, nclass) = zip(*tasks)
        root_t = torch.tensor(np.stack(root), dtype=torch.float32)
        events_t = torch.tensor(np.stack(events), dtype=torch.float32)
        ev_mask_t = torch.tensor(np.stack(ev_mask), dtype=torch.bool)
        tab_mask_t = torch.tensor(np.stack(tab_mask), dtype=torch.bool)
        lag_t = torch.tensor(np.stack(lag), dtype=torch.float32)
        time_t = torch.tensor(np.stack(time_feat), dtype=torch.float32)
        target_np = np.stack(target).astype(np.float32)
        y_class_np = np.stack(y_class).astype(np.int64)
        context_target = target_np.copy()
        query_mask = np.zeros_like(context_target, dtype=bool)
        query_mask[:, -1] = True
        context_target[:, -1] = 0.0
        return MultiTableBatch(
            root=root_t,
            events=events_t,
            event_mask=ev_mask_t,
            table_mask=tab_mask_t,
            context_target=torch.tensor(context_target, dtype=torch.float32),
            lag_target=lag_t,
            time=time_t,
            query_mask=torch.tensor(query_mask, dtype=torch.bool),
            task_type=torch.tensor(np.array(ttype), dtype=torch.long),
            num_classes=torch.tensor(np.array(nclass), dtype=torch.long),
            y=torch.tensor(target_np[:, -1], dtype=torch.float32),
            y_class=torch.tensor(y_class_np[:, -1], dtype=torch.long),
        )

    def _task(self):
        n = self.samples
        N = self.num_event_tables
        R = self.rows_per_event
        latent = self.rng.normal(size=(n, 3)).astype(np.float32)
        root = self.rng.normal(size=(n, self.root_cols)).astype(np.float32)
        root[:, :3] = 0.65 * root[:, :3] + latent
        events = self.rng.normal(size=(n, N, R, self.event_cols)).astype(np.float32)
        events[:, :, :, :3] += latent[:, None, None, :]
        # Per-table identity offset so tables are distinguishable in signal.
        table_offset = self.rng.normal(size=(N, 3)).astype(np.float32) * 0.5
        events[:, :, :, :3] += table_offset[None, :, None, :]
        if self.noise_cols > 0:
            nc = min(self.noise_cols, self.event_cols)
            events[:, :, :, -nc:] = self.rng.normal(size=(n, N, R, nc))
            rnc = min(self.noise_cols, self.root_cols)
            root[:, -rnc:] = self.rng.normal(size=(n, rnc))

        anchor = np.sort(self.rng.uniform(0.05, 1.0, size=n)).astype(np.float32)
        local = np.zeros(n, dtype=np.float32)
        local[-max(1, n // 4):] = 1.0
        self.rng.shuffle(local[:-1])
        local[-1] = 1.0

        ev_mask = np.ones((n, N, R), dtype=bool)
        if self.link_dropout > 0:
            ev_mask &= self.rng.random(size=ev_mask.shape) > self.link_dropout
            ev_mask[:, :, 0] = True
        # Table availability (schema dropout): some event tables missing for the whole task.
        tab_avail = np.ones(N + 1, dtype=bool)
        if self.table_dropout > 0:
            tab_avail[1:] = self.rng.random(size=N) > self.table_dropout
            if not tab_avail[1:].any():
                tab_avail[1] = True
        table_mask = np.broadcast_to(tab_avail, (n, N + 1)).copy()
        for t in range(N):
            if not tab_avail[t + 1]:
                ev_mask[:, t] = False
                events[:, t] = 0.0

        # Label: cross-table interaction over a RANDOM SUBSET of AVAILABLE tables.
        avail = [t for t in range(N) if tab_avail[t + 1]]
        k = self.rng.integers(1, len(avail) + 1)
        chosen = self.rng.choice(avail, size=int(k), replace=False)
        mechanism = int(self.rng.integers(0, 4))
        score = np.zeros(n, dtype=np.float32)
        for t in chosen:
            col = int(self.rng.integers(0, max(1, self.event_cols - self.noise_cols)))
            masked = events[:, t, :, col] * ev_mask[:, t]
            denom = ev_mask[:, t].sum(axis=1).clip(min=1)
            if mechanism == 0:
                score += masked.sum(axis=1) / denom                       # mean
            elif mechanism == 1:
                score += np.maximum(masked, -4).max(axis=1)               # max
            elif mechanism == 2:
                score += (masked > 0.3).sum(axis=1).astype(np.float32) / denom  # rate
            else:
                gate = events[:, t, :, 0] + root[:, None, 0] > 0.0        # root-gated mean
                gm = (events[:, t, :, col] * (ev_mask[:, t] & gate)).sum(axis=1) / (ev_mask[:, t] & gate).sum(axis=1).clip(min=1)
                score += gm
        score = (score / max(1, len(chosen))).astype(np.float32)
        # Per-task random sign: the feature->label direction is arbitrary, so the ONLY
        # way to orient predictions is to read the visible CONTEXT labels (true ICL).
        # Without this the model memorizes a fixed feature->label sign that inverts on
        # real data whose feature semantics differ.
        if self.rng.random() < 0.5:
            score = -score
        score += 0.3 * root[:, 0] + 0.2 * anchor + 0.15 * self.rng.normal(size=n).astype(np.float32)
        # Inject same-entity label-history autocorrelation so the model learns to
        # propagate visible context labels (the mechanism that sets direction at eval).
        score = self._inject_label_history(score, anchor)

        # Root-shortcut dropout: after the label is fixed, replace the query root's static
        # features with noise for a fraction of tasks so the model cannot solve the task from
        # query root features alone and must read relational context (adversarial to shortcuts).
        if self.root_shortcut_dropout > 0 and self.rng.random() < self.root_shortcut_dropout:
            root[-1, :] = self.rng.normal(size=self.root_cols).astype(np.float32)

        ttype = self._choose_task_type()
        target, y_class, nclass = self._target_from_score(score, ttype)
        lag = self._lagged(target, anchor)
        time_feat = np.stack([anchor, local], axis=1).astype(np.float32)
        # missingness
        if self.missing_prob > 0:
            root[self.rng.random(root.shape) < self.missing_prob] = 0.0
            events[self.rng.random(events.shape) < self.missing_prob] = 0.0
        events[~ev_mask] = 0.0
        # Column-order invariance: permute feature columns per task so the model
        # cannot rely on fixed column indices (real schemas have arbitrary column
        # order/semantics). Validated as transfer-critical in the 3-table effort.
        rp = self.rng.permutation(self.root_cols)
        root = root[:, rp]
        cp = self.rng.permutation(self.event_cols)
        events = events[:, :, :, cp]
        return root, events, ev_mask, table_mask, lag, time_feat, target, y_class, ttype, nclass

    def _choose_task_type(self) -> int:
        if self.task_type == "mixed":
            return int(self.rng.integers(0, 3))
        return {"binary": 0, "regression": 1, "multiclass": 2}[self.task_type]

    def _target_from_score(self, score, ttype):
        if ttype == 0:
            qa = float(np.clip(self.binary_quantile_min, 0.01, 0.99))
            qb = float(np.clip(self.binary_quantile_max, qa, 0.99))
            q = qa if qa == qb else float(self.rng.uniform(qa, qb))
            thr = np.quantile(score[:-1], q)
            return (score > thr).astype(np.float32), np.full(len(score), -1, np.int64), 2
        if ttype == 1:
            mean = score[:-1].mean(); std = score[:-1].std().clip(min=1e-3)
            return ((score - mean) / std).astype(np.float32), np.full(len(score), -1, np.int64), 1
        nclass = int(self.rng.integers(3, self.max_classes + 1))
        bins = np.quantile(score[:-1], np.linspace(0, 1, nclass + 1)[1:-1])
        classes = np.digitize(score, bins).astype(np.int64)
        return (classes / max(1, nclass - 1)).astype(np.float32), classes, nclass

    def _inject_label_history(self, score, anchor):
        """Blend the feature score with a standardized autoregressive component so a
        context example's label is predictable from earlier same-task labels. This
        teaches label propagation, which is how the model recovers the (randomized)
        label direction from visible context at inference.
        """
        n = len(score)
        std = (score - score.mean()) / (score.std() + 1e-6)
        ar = np.zeros(n, dtype=np.float32)
        for i in range(1, n):
            ar[i] = 0.6 * std[i - 1] + 0.4 * std[max(0, i - 2)]
        # history_label_frac (when set) fixes a higher autoregressive weight, forcing the
        # model to propagate visible context labels — the ingredient that first beat the
        # strict root-only baseline in the 3-table effort. 0 keeps the old random blend.
        frac = self.history_label_frac if self.history_label_frac > 0 else float(self.rng.uniform(0.3, 0.7))
        return ((1 - frac) * std + frac * ar).astype(np.float32)

    def _lagged(self, y, anchor):
        lag = np.zeros((len(y), self.lag_steps), dtype=np.float32)
        for i in range(len(y)):
            for l in range(self.lag_steps):
                j = i - l - 1
                lag[i, l] = y[j] if j >= 0 and anchor[j] < anchor[i] else 0.0
        return lag


def multitable_task_loss(output: torch.Tensor, batch: MultiTableBatch) -> torch.Tensor:
    import torch.nn.functional as F
    output = output.float()
    losses = []
    binary = batch.task_type == 0
    if binary.any():
        losses.append(F.binary_cross_entropy_with_logits(output[binary, 0], batch.y[binary].float()))
    reg = batch.task_type == 1
    if reg.any():
        losses.append(F.mse_loss(output[reg, 1], batch.y[reg]))
    mc = batch.task_type == 2
    if mc.any():
        cl = output[mc, 2:]
        ids = torch.arange(cl.shape[1], device=cl.device).unsqueeze(0)
        cl = cl.masked_fill(ids >= batch.num_classes[mc].unsqueeze(1), -1e4)
        losses.append(F.cross_entropy(cl, batch.y_class[mc]))
    if not losses:
        return output.sum() * 0.0
    return torch.stack(losses).mean()
