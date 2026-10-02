from __future__ import annotations

import math

import torch
from torch import nn


class RotaryPositionalEmbedding(nn.Module):
    """RoPE for sequence position awareness in table rows and cross-sample ordering."""

    def __init__(self, dim: int, max_len: int = 2048) -> None:
        super().__init__()
        inv_freq = 1.0 / (10000 ** (torch.arange(0, dim, 2).float() / dim))
        self.register_buffer("inv_freq", inv_freq)
        self.max_len = max_len

    def forward(self, x: torch.Tensor) -> torch.Tensor:
        seq_len = x.shape[-2]
        t = torch.arange(seq_len, device=x.device, dtype=self.inv_freq.dtype)
        freqs = torch.outer(t, self.inv_freq)
        emb = torch.cat((freqs, freqs), dim=-1)
        cos = emb.cos()
        sin = emb.sin()
        ndim = x.ndim
        for _ in range(ndim - 2):
            cos = cos.unsqueeze(0)
            sin = sin.unsqueeze(0)
        return x * cos + rotate_half(x) * sin


def rotate_half(x: torch.Tensor) -> torch.Tensor:
    x1, x2 = x.chunk(2, dim=-1)
    return torch.cat((-x2, x1), dim=-1)


class ImprovedAttnBlock(nn.Module):
    """Pre-norm attention block with optional RoPE and standard FFN for stability."""

    def __init__(self, d_model: int, heads: int, dropout: float, use_rope: bool = False) -> None:
        super().__init__()
        self.norm1 = nn.LayerNorm(d_model)
        self.attn = nn.MultiheadAttention(d_model, heads, dropout=dropout, batch_first=True)
        self.norm2 = nn.LayerNorm(d_model)
        self.ffn = nn.Sequential(
            nn.Linear(d_model, 4 * d_model),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(4 * d_model, d_model),
        )
        self.drop = nn.Dropout(dropout)
        self.rope = RotaryPositionalEmbedding(d_model) if use_rope else None
        self.use_rope = use_rope

    def forward(self, x: torch.Tensor, key_padding_mask: torch.Tensor | None = None) -> torch.Tensor:
        h = self.norm1(x)
        if self.use_rope:
            h = self.rope(h)
        if x.device.type == "cuda" and torch.is_autocast_enabled("cuda"):
            with torch.amp.autocast("cuda", enabled=False):
                attn_in = h.float()
                h, _ = self.attn(attn_in, attn_in, attn_in, key_padding_mask=key_padding_mask, need_weights=False)
            h = h.to(x.dtype)
        else:
            h, _ = self.attn(h, h, h, key_padding_mask=key_padding_mask, need_weights=False)
        x = x + self.drop(h)
        x = x + self.drop(self.ffn(self.norm2(x)))
        return x


class ImprovedTypedGraphBlock(nn.Module):
    """Pairwise typed attention over table tokens.

    The graph has only a few table nodes, so an explicit attention
    implementation is clearer than trying to shoehorn relation-specific biases
    into ``nn.MultiheadAttention``. Each target/source table pair gets a learned
    per-head logit bias and a learned value embedding.
    """

    def __init__(self, num_tables: int, d_model: int, heads: int, dropout: float) -> None:
        super().__init__()
        if d_model % heads != 0:
            raise ValueError(f"d_model ({d_model}) must be divisible by heads ({heads})")
        self.num_tables = num_tables
        self.heads = heads
        self.head_dim = d_model // heads
        self.norm1 = nn.LayerNorm(d_model)
        self.qkv = nn.Linear(d_model, 3 * d_model)
        self.out = nn.Linear(d_model, d_model)
        self.edge_bias = nn.Parameter(torch.zeros(num_tables, num_tables, heads))
        self.edge_value = nn.Parameter(torch.randn(num_tables, num_tables, d_model) * 0.01)
        self.norm2 = nn.LayerNorm(d_model)
        self.ffn = nn.Sequential(
            nn.Linear(d_model, 4 * d_model),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(4 * d_model, d_model),
        )
        self.drop = nn.Dropout(dropout)

    def forward(self, x: torch.Tensor, key_padding_mask: torch.Tensor | None = None) -> torch.Tensor:
        n, num_t, d = x.shape
        h = self.norm1(x)
        compute_fp32 = x.device.type == "cuda" and torch.is_autocast_enabled("cuda")
        if compute_fp32:
            h_in = h.float()
            qkv_weight = self.qkv.weight.float()
            qkv_bias = self.qkv.bias.float() if self.qkv.bias is not None else None
            out_weight = self.out.weight.float()
            out_bias = self.out.bias.float() if self.out.bias is not None else None
            edge_bias = self.edge_bias[:num_t, :num_t].float()
            edge_value = self.edge_value[:num_t, :num_t].float()
        else:
            h_in = h
            qkv_weight = self.qkv.weight
            qkv_bias = self.qkv.bias
            out_weight = self.out.weight
            out_bias = self.out.bias
            edge_bias = self.edge_bias[:num_t, :num_t]
            edge_value = self.edge_value[:num_t, :num_t]

        qkv = torch.nn.functional.linear(h_in, qkv_weight, qkv_bias)
        q, k, v = qkv.chunk(3, dim=-1)
        q = q.reshape(n, num_t, self.heads, self.head_dim).transpose(1, 2)
        k = k.reshape(n, num_t, self.heads, self.head_dim).transpose(1, 2)
        v = v.reshape(n, num_t, self.heads, self.head_dim).transpose(1, 2)
        logits = torch.matmul(q, k.transpose(-1, -2)) / math.sqrt(self.head_dim)
        logits = logits + edge_bias.permute(2, 0, 1).unsqueeze(0)
        if key_padding_mask is not None:
            logits = logits.masked_fill(key_padding_mask[:, None, None, :], -1.0e4)
        attn = torch.softmax(logits.float(), dim=-1).to(v.dtype)
        rel_v = edge_value.reshape(num_t, num_t, self.heads, self.head_dim).permute(2, 0, 1, 3)
        messages = v[:, :, None, :, :] + rel_v[None, :, :, :, :]
        graph_h = (attn[..., None] * messages).sum(dim=-2)
        graph_h = graph_h.transpose(1, 2).reshape(n, num_t, d)
        graph_h = torch.nn.functional.linear(graph_h, out_weight, out_bias)
        graph_h = graph_h.to(x.dtype)
        x = x + self.drop(graph_h)
        x = x + self.drop(self.ffn(self.norm2(x)))
        return x


class ImprovedTableEncoder(nn.Module):
    """Table encoder with column-type awareness, RoPE on rows, and attention pooling.

    Alternates column-wise and row-wise attention. Column embeddings are typed
    (feature, time, identity) and learn a per-column position. Row attention
    uses RoPE for order sensitivity, and final pooling uses attention weighting
    rather than simple averaging when masks are absent.
    """

    def __init__(
        self,
        num_cols: int,
        d_model: int,
        heads: int,
        layers: int,
        dropout: float,
        column_roles: torch.Tensor | None = None,
    ) -> None:
        super().__init__()
        self.col_pos = nn.Parameter(torch.randn(num_cols, d_model) * 0.02)
        self.value = nn.Linear(1, d_model)
        if column_roles is not None:
            if column_roles.numel() != num_cols:
                raise ValueError(f"column_roles length {column_roles.numel()} does not match num_cols {num_cols}")
            self.register_buffer("column_roles", column_roles.long(), persistent=False)
            self.role_emb = nn.Embedding(5, d_model)
        else:
            self.column_roles = None
            self.role_emb = None
        self.col_blocks = nn.ModuleList([ImprovedAttnBlock(d_model, heads, dropout) for _ in range(layers)])
        self.row_blocks = nn.ModuleList([ImprovedAttnBlock(d_model, heads, dropout, use_rope=True) for _ in range(layers)])
        self.out = nn.LayerNorm(d_model)
        self.attn_pool_query = nn.Parameter(torch.randn(1, 1, d_model) * 0.02)

    def forward(
        self,
        x: torch.Tensor,
        row_mask: torch.Tensor | None = None,
    ) -> torch.Tensor:
        b, s, r, c = x.shape
        h = self.value(x.reshape(b * s * r, c, 1))
        h = h + self.col_pos[:c].unsqueeze(0)
        if self.role_emb is not None and self.column_roles is not None:
            h = h + self.role_emb(self.column_roles[:c]).unsqueeze(0)
        row_pad = None
        if row_mask is not None:
            row_pad = ~row_mask.reshape(b * s, r)
            if row_pad.all(dim=1).any():
                row_pad = row_pad.clone()
                row_pad[row_pad.all(dim=1), 0] = False

        for col_block, row_block in zip(self.col_blocks, self.row_blocks):
            h = col_block(h)
            row_tokens = h.mean(dim=1).reshape(b * s, r, -1)
            row_tokens = row_block(row_tokens, key_padding_mask=row_pad)
            h = h + row_tokens.reshape(b * s * r, 1, -1)

        row_emb = h.mean(dim=1).reshape(b, s, r, -1)
        query = self.attn_pool_query.expand(b * s, 1, -1)
        raw_weights = torch.matmul(query, row_emb.reshape(b * s, r, -1).transpose(-1, -2))
        scale = math.sqrt(row_emb.shape[-1])
        raw_weights = (raw_weights / scale).clamp(-10.0, 10.0)
        if row_mask is not None:
            flat_mask = row_mask.reshape(b * s, r)
            safe_mask = flat_mask
            if safe_mask.logical_not().all(dim=1).any():
                safe_mask = safe_mask.clone()
                safe_mask[safe_mask.logical_not().all(dim=1), 0] = True
            raw_weights = raw_weights.squeeze(1).masked_fill(~safe_mask, -1.0e4).unsqueeze(1)
        weights = torch.softmax(raw_weights.float(), dim=-1).to(row_emb.dtype)
        if row_mask is not None:
            flat_mask_f = row_mask.reshape(b * s, r).to(weights.dtype).unsqueeze(1)
            weights = weights * flat_mask_f
            weights = weights / weights.sum(dim=-1, keepdim=True).clamp_min(torch.finfo(weights.dtype).eps)
        pooled = (weights @ row_emb.reshape(b * s, r, -1)).squeeze(1).reshape(b, s, -1)
        return self.out(pooled)


class ImprovedGraphCrossSampleEncoder(nn.Module):
    """Enhanced graph encoder with typed edges and learned relational structure.

    Alternates graph-level (cross-table) and sample-level (cross-context) attention.
    Edge type embeddings distinguish root↔child, root↔aux, child↔aux interactions.
    """

    def __init__(
        self,
        num_tables: int,
        d_model: int,
        heads: int,
        layers: int,
        dropout: float,
        use_typed_pair_attention: bool = True,
    ) -> None:
        super().__init__()
        self.table_type_emb = nn.Parameter(torch.randn(num_tables, d_model) * 0.02)
        self.use_typed_pair_attention = use_typed_pair_attention
        if use_typed_pair_attention:
            self.typed_graph_blocks = nn.ModuleList(
                [ImprovedTypedGraphBlock(num_tables, d_model, heads, dropout) for _ in range(layers)]
            )
        else:
            self.edge_types = nn.Parameter(torch.randn(num_tables, num_tables, d_model) * 0.01)
            self.graph_blocks = nn.ModuleList([ImprovedAttnBlock(d_model, heads, dropout) for _ in range(layers)])
        self.sample_blocks = nn.ModuleList([ImprovedAttnBlock(d_model, heads, dropout, use_rope=True) for _ in range(layers)])

    def forward(self, tables: list[torch.Tensor], table_mask: torch.Tensor | None = None) -> torch.Tensor:
        b, s, d = tables[0].shape
        num_t = len(tables)
        table = torch.stack(tables, dim=2)
        table = table + self.table_type_emb[:num_t].view(1, 1, num_t, d)

        table_pad = None
        if table_mask is not None:
            table_pad = ~table_mask.reshape(b * s, num_t)
            if table_pad.all(dim=1).any():
                table_pad = table_pad.clone()
                table_pad[table_pad.all(dim=1), 0] = False

        graph_blocks = self.typed_graph_blocks if self.use_typed_pair_attention else self.graph_blocks
        for idx, (graph_block, sample_block) in enumerate(zip(graph_blocks, self.sample_blocks)):
            graph_in = table.reshape(b * s, num_t, d)
            if not self.use_typed_pair_attention:
                edge_bias = self.edge_types[:num_t, :num_t].mean().reshape(()) * 0.1
                graph_in = graph_in + edge_bias
            graph_out = graph_block(graph_in, key_padding_mask=table_pad)
            table = graph_out.reshape(b, s, num_t, d)

            if table_mask is None:
                sample = table.mean(dim=2)
            else:
                weights = table_mask.float().unsqueeze(-1)
                sample = (table * weights).sum(dim=2) / weights.sum(dim=2).clamp_min(1.0)
            sample = sample_block(sample)
            table = table + sample.unsqueeze(2)

        if table_mask is None:
            return table.mean(dim=2)
        weights = table_mask.float().unsqueeze(-1)
        return (table * weights).sum(dim=2) / weights.sum(dim=2).clamp_min(1.0)


class RowBridgeAttention(nn.Module):
    """Row-level child<->aux attention before table-level pooling.

    Table pooling is efficient, but it can erase join-like evidence where a
    child row only matters together with a matching auxiliary row. This module
    adds a generic, learned row-to-row interaction summary that is fused into
    the sample embedding after graph aggregation.
    """

    def __init__(
        self,
        root_cols: int,
        child_cols: int,
        aux_cols: int,
        task_cols: int,
        d_model: int,
        dropout: float,
    ) -> None:
        super().__init__()
        self.root_proj = nn.Linear(root_cols + task_cols, d_model)
        self.child_q = nn.Linear(child_cols + task_cols, d_model)
        self.child_k = nn.Linear(child_cols + task_cols, d_model)
        self.child_v = nn.Linear(child_cols + task_cols, d_model)
        self.aux_q = nn.Linear(aux_cols + task_cols, d_model)
        self.aux_k = nn.Linear(aux_cols + task_cols, d_model)
        self.aux_v = nn.Linear(aux_cols + task_cols, d_model)
        self.fuse = nn.Sequential(
            nn.LayerNorm(3 * d_model),
            nn.Linear(3 * d_model, d_model),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(d_model, d_model),
        )
        self.out = nn.LayerNorm(d_model)

    def forward(
        self,
        root: torch.Tensor,
        child: torch.Tensor,
        aux: torch.Tensor,
        child_mask: torch.Tensor,
        aux_mask: torch.Tensor,
        table_mask: torch.Tensor,
    ) -> torch.Tensor:
        b, s, child_rows, _ = child.shape
        aux_rows = aux.shape[2]
        child_flat = child.reshape(b * s, child_rows, -1)
        aux_flat = aux.reshape(b * s, aux_rows, -1)
        child_mask_flat = child_mask.reshape(b * s, child_rows)
        aux_mask_flat = aux_mask.reshape(b * s, aux_rows)
        table_flat = table_mask.reshape(b * s, -1)

        root_h = self.root_proj(root.reshape(b * s, -1))
        child_q = self.child_q(child_flat)
        child_k = self.child_k(child_flat)
        child_v = self.child_v(child_flat)
        aux_q = self.aux_q(aux_flat)
        aux_k = self.aux_k(aux_flat)
        aux_v = self.aux_v(aux_flat)

        child_to_aux = self._attend(child_q, aux_k, aux_v, aux_mask_flat, child_mask_flat)
        aux_to_child = self._attend(aux_q, child_k, child_v, child_mask_flat, aux_mask_flat)
        child_summary = self._masked_mean(child_q * child_to_aux, child_mask_flat)
        aux_summary = self._masked_mean(aux_q * aux_to_child, aux_mask_flat)

        table_available = (table_flat[:, 1] & table_flat[:, 2]).to(child_summary.dtype).unsqueeze(-1)
        fused = self.fuse(torch.cat([root_h, child_summary, aux_summary], dim=-1))
        fused = fused * table_available
        return self.out(fused.reshape(b, s, -1))

    @staticmethod
    def _attend(
        query: torch.Tensor,
        key: torch.Tensor,
        value: torch.Tensor,
        key_mask: torch.Tensor,
        query_mask: torch.Tensor,
    ) -> torch.Tensor:
        logits = torch.matmul(query.float(), key.float().transpose(-1, -2)) / math.sqrt(query.shape[-1])
        safe_key_mask = key_mask
        if safe_key_mask.logical_not().all(dim=1).any():
            safe_key_mask = safe_key_mask.clone()
            safe_key_mask[safe_key_mask.logical_not().all(dim=1), 0] = True
        logits = logits.masked_fill(~safe_key_mask[:, None, :], -1.0e4)
        weights = torch.softmax(logits, dim=-1).to(value.dtype)
        attended = torch.matmul(weights, value)
        return attended * query_mask.to(attended.dtype).unsqueeze(-1)

    @staticmethod
    def _masked_mean(values: torch.Tensor, mask: torch.Tensor) -> torch.Tensor:
        weights = mask.to(values.dtype).unsqueeze(-1)
        return (values * weights).sum(dim=1) / weights.sum(dim=1).clamp_min(1.0)


class RelationalFoundationModel(nn.Module):
    """KumoRFM-style relational foundation model with enhanced architecture.

    Key improvements over the baseline:
    - Gated FFN with sigmoid activation for better gradient flow
    - RoPE on row and sample attention for position awareness
    - Column-type embeddings distinguishing feature/time/identity columns
    - Attention-weighted table pooling (not just mean)
    - Edge-type embeddings for typed cross-table interactions
    - Deeper readout network with residual connections
    - Layer normalization at more locations for training stability
    """

    def __init__(
        self,
        root_cols: int,
        child_cols: int,
        aux_cols: int = 6,
        lag_steps: int = 4,
        max_classes: int = 8,
        d_model: int = 256,
        heads: int = 8,
        table_layers: int = 3,
        graph_layers: int = 3,
        dropout: float = 0.1,
        use_improved: bool = True,
        typed_graph_attention: bool = True,
        typed_task_conditioning: bool = False,
        row_bridge_attention: bool = False,
    ) -> None:
        super().__init__()
        task_cols = 2 + lag_steps + 2
        encoder_cls = ImprovedTableEncoder if use_improved else TableEncoder
        graph_cls = ImprovedGraphCrossSampleEncoder if use_improved else GraphCrossSampleEncoder

        root_roles = task_column_roles(root_cols, lag_steps) if typed_task_conditioning else None
        child_roles = task_column_roles(child_cols, lag_steps) if typed_task_conditioning else None
        aux_roles = task_column_roles(aux_cols, lag_steps) if typed_task_conditioning else None

        self.root_encoder = encoder_cls(root_cols + task_cols, d_model, heads, table_layers, dropout, root_roles)
        self.child_encoder = encoder_cls(child_cols + task_cols, d_model, heads, table_layers, dropout, child_roles)
        self.aux_encoder = encoder_cls(aux_cols + task_cols, d_model, heads, table_layers, dropout, aux_roles)
        if use_improved:
            self.graph_encoder = graph_cls(3, d_model, heads, graph_layers, dropout, typed_graph_attention)
        else:
            self.graph_encoder = graph_cls(3, d_model, heads, graph_layers, dropout)
        self.row_bridge_attention = row_bridge_attention
        self.row_bridge = (
            RowBridgeAttention(root_cols, child_cols, aux_cols, task_cols, d_model, dropout)
            if row_bridge_attention
            else None
        )
        self.max_classes = max_classes

        readout_dim = 4 * d_model
        self.readout = nn.Sequential(
            nn.LayerNorm(d_model),
            nn.Linear(d_model, readout_dim),
            nn.GELU(),
            nn.Linear(readout_dim, d_model),
            nn.GELU(),
            nn.LayerNorm(d_model),
            nn.Linear(d_model, 2 + max_classes),
        )

    def forward(
        self,
        root: torch.Tensor,
        child: torch.Tensor,
        aux: torch.Tensor,
        child_mask: torch.Tensor,
        aux_mask: torch.Tensor,
        table_mask: torch.Tensor,
        context_target: torch.Tensor,
        lag_target: torch.Tensor,
        time: torch.Tensor,
        query_mask: torch.Tensor,
        return_embedding: bool = False,
    ) -> torch.Tensor | tuple[torch.Tensor, torch.Tensor]:
        emb = self.encode(root, child, aux, child_mask, aux_mask, table_mask, context_target, lag_target, time, query_mask)
        logits = self.readout(emb)
        if return_embedding:
            return logits, emb
        return logits

    def encode(
        self,
        root: torch.Tensor,
        child: torch.Tensor,
        aux: torch.Tensor,
        child_mask: torch.Tensor,
        aux_mask: torch.Tensor,
        table_mask: torch.Tensor,
        context_target: torch.Tensor,
        lag_target: torch.Tensor,
        time: torch.Tensor,
        query_mask: torch.Tensor,
    ) -> torch.Tensor:
        target_feat = context_target.unsqueeze(-1)
        visible_feat = (~query_mask).float().unsqueeze(-1)
        task_feat = torch.cat([target_feat, visible_feat, lag_target, time], dim=-1)

        root_in = torch.cat([root, task_feat], dim=-1).unsqueeze(2)
        child_extra = task_feat[:, :, None, :].expand(-1, -1, child.shape[2], -1)
        child_in = torch.cat([child, child_extra], dim=-1)
        aux_extra = task_feat[:, :, None, :].expand(-1, -1, aux.shape[2], -1)
        aux_in = torch.cat([aux, aux_extra], dim=-1)

        root_emb = self.root_encoder(root_in)
        child_emb = self.child_encoder(child_in, row_mask=child_mask)
        aux_emb = self.aux_encoder(aux_in, row_mask=aux_mask)
        sample_emb = self.graph_encoder([root_emb, child_emb, aux_emb], table_mask=table_mask)
        if self.row_bridge is not None:
            sample_emb = sample_emb + self.row_bridge(root_in.squeeze(2), child_in, aux_in, child_mask, aux_mask, table_mask)
        query_index = query_mask.float().argmax(dim=1)
        b = root.shape[0]
        return sample_emb[torch.arange(b, device=root.device), query_index]


def task_column_roles(feature_cols: int, lag_steps: int) -> torch.Tensor:
    """Column roles for feature plus task-conditioning columns.

    Roles: 0=raw feature, 1=visible target value, 2=target visibility/query
    marker, 3=lagged target, 4=time/local-context feature.
    """

    return torch.tensor([0] * feature_cols + [1, 2] + [3] * lag_steps + [4, 4], dtype=torch.long)


class TableEncoder(nn.Module):
    """Baseline table encoder kept for backward compatibility.

    Alternating column-wise and row-wise attention over a table.
    """

    def __init__(
        self,
        num_cols: int,
        d_model: int,
        heads: int,
        layers: int,
        dropout: float,
        column_roles: torch.Tensor | None = None,
    ) -> None:
        super().__init__()
        self.col_id = nn.Parameter(torch.randn(num_cols, d_model) * 0.02)
        self.value = nn.Linear(1, d_model)
        if column_roles is not None:
            if column_roles.numel() != num_cols:
                raise ValueError(f"column_roles length {column_roles.numel()} does not match num_cols {num_cols}")
            self.register_buffer("column_roles", column_roles.long(), persistent=False)
            self.role_emb = nn.Embedding(5, d_model)
        else:
            self.column_roles = None
            self.role_emb = None
        self.col_blocks = nn.ModuleList([AttnBlock(d_model, heads, dropout) for _ in range(layers)])
        self.row_blocks = nn.ModuleList([AttnBlock(d_model, heads, dropout) for _ in range(layers)])
        self.out = nn.LayerNorm(d_model)

    def forward(self, x: torch.Tensor, row_mask: torch.Tensor | None = None) -> torch.Tensor:
        b, s, r, c = x.shape
        h = self.value(x.reshape(b * s * r, c, 1)) + self.col_id[:c].unsqueeze(0)
        if self.role_emb is not None and self.column_roles is not None:
            h = h + self.role_emb(self.column_roles[:c]).unsqueeze(0)
        row_pad = None
        if row_mask is not None:
            row_pad = ~row_mask.reshape(b * s, r)
            if row_pad.all(dim=1).any():
                row_pad = row_pad.clone()
                row_pad[row_pad.all(dim=1), 0] = False

        for col_block, row_block in zip(self.col_blocks, self.row_blocks):
            h = col_block(h)
            row_tokens = h.mean(dim=1).reshape(b * s, r, -1)
            row_tokens = row_block(row_tokens, key_padding_mask=row_pad)
            h = h + row_tokens.reshape(b * s * r, 1, -1)

        row_emb = h.mean(dim=1).reshape(b, s, r, -1)
        if row_mask is None:
            pooled = row_emb.mean(dim=2)
        else:
            weights = row_mask.float()
            pooled = (row_emb * weights.unsqueeze(-1)).sum(dim=2) / weights.sum(dim=2, keepdims=True).clamp_min(1.0)
        return self.out(pooled)


class AttnBlock(nn.Module):
    def __init__(self, d_model: int, heads: int, dropout: float) -> None:
        super().__init__()
        self.norm1 = nn.LayerNorm(d_model)
        self.attn = nn.MultiheadAttention(d_model, heads, dropout=dropout, batch_first=True)
        self.norm2 = nn.LayerNorm(d_model)
        self.ffn = nn.Sequential(
            nn.Linear(d_model, 4 * d_model),
            nn.GELU(),
            nn.Dropout(dropout),
            nn.Linear(4 * d_model, d_model),
        )
        self.drop = nn.Dropout(dropout)

    def forward(self, x: torch.Tensor, key_padding_mask: torch.Tensor | None = None) -> torch.Tensor:
        h = self.norm1(x)
        h, _ = self.attn(h, h, h, key_padding_mask=key_padding_mask, need_weights=False)
        x = x + self.drop(h)
        x = x + self.drop(self.ffn(self.norm2(x)))
        return x


class GraphCrossSampleEncoder(nn.Module):
    """Baseline encoder kept for backward compatibility."""

    def __init__(self, num_tables: int, d_model: int, heads: int, layers: int, dropout: float) -> None:
        super().__init__()
        self.table_type = nn.Parameter(torch.randn(num_tables, d_model) * 0.02)
        self.graph_blocks = nn.ModuleList([AttnBlock(d_model, heads, dropout) for _ in range(layers)])
        self.sample_blocks = nn.ModuleList([AttnBlock(d_model, heads, dropout) for _ in range(layers)])

    def forward(self, tables: list[torch.Tensor], table_mask: torch.Tensor | None = None) -> torch.Tensor:
        b, s, d = tables[0].shape
        table = torch.stack(tables, dim=2)
        table = table + self.table_type[: len(tables)].view(1, 1, len(tables), d)
        table_pad = None
        if table_mask is not None:
            table_pad = ~table_mask.reshape(b * s, len(tables))
            if table_pad.all(dim=1).any():
                table_pad = table_pad.clone()
                table_pad[table_pad.all(dim=1), 0] = False
        for graph_block, sample_block in zip(self.graph_blocks, self.sample_blocks):
            num_tables = table.shape[2]
            table = graph_block(table.reshape(b * s, num_tables, d), key_padding_mask=table_pad).reshape(b, s, num_tables, d)
            if table_mask is None:
                sample = table.mean(dim=2)
            else:
                weights = table_mask.float().unsqueeze(-1)
                sample = (table * weights).sum(dim=2) / weights.sum(dim=2).clamp_min(1.0)
            sample = sample_block(sample)
            table = table + sample.unsqueeze(2)
        if table_mask is None:
            return table.mean(dim=2)
        weights = table_mask.float().unsqueeze(-1)
        return (table * weights).sum(dim=2) / weights.sum(dim=2).clamp_min(1.0)
