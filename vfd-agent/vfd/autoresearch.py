"""P4: a Karpathy-style auto-research loop over the harness configuration, with the overfitting guards of paper 1.

  1. Questions are split once, before the loop: SEARCH (where changes are tried), ACCEPT (a slice inside development a
     change must not hurt), HELDOUT (scored once at the very end, for the initial and the final configuration only).
  2. A researcher LLM reads the current configuration, the allowed knobs and the full history of attempts, and proposes
     one change with a one-line hypothesis.
  3. The change is run on SEARCH and ACCEPT. It is kept only if the paired per-question gain on SEARCH exceeds
     `z` standard errors AND the gain on ACCEPT is not negative.
  4. Fixed budget of attempts; every attempt is logged with its scores and the decision.
`evaluate(config, qids) -> {qid: score}` is supplied by the caller (e.g. the SQL harness: 1 if correct).
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any, Callable

import numpy as np

PROMPT = """You improve an agent configuration by running experiments. Each experiment changes ONE thing.
Allowed knobs and values:
{space}
Current best configuration (score {best:.4f} on the search questions):
{config}
History of attempts (newest last):
{history}
Propose the next experiment. Return only JSON: {{"knob": ..., "value": ..., "hypothesis": "..."}}. Do not repeat an
attempt that was already tried with the same value."""


@dataclass
class AutoResearch:
    evaluate: Callable[[dict, list], dict]
    space: dict[str, list]
    researcher: Any                       # vfd.llm.Chat
    budget: int = 20
    z: float = 1.0
    seed: int = 0
    history: list = field(default_factory=list)

    def split(self, qids: list, search: float = 0.6, accept: float = 0.15) -> tuple[list, list, list]:
        rng = np.random.default_rng(self.seed)
        q = list(rng.permutation(qids))
        a, b = int(len(q) * search), int(len(q) * (search + accept))
        return q[:a], q[a:b], q[b:]

    @staticmethod
    def paired(new: dict, old: dict, qids: list) -> tuple[float, float]:
        d = np.array([new[q] - old[q] for q in qids if q in new and q in old], float)
        if len(d) < 2:
            return float(d.mean()) if len(d) else 0.0, float("inf")
        return float(d.mean()), float(d.std(ddof=1) / np.sqrt(len(d)))

    def propose(self, best_cfg: dict, best: float) -> dict | None:
        hist = "\n".join(f"- {h['knob']}={h['value']}: search {h['search_delta']:+.4f} (se {h['search_se']:.4f}), "
                         f"accept {h['accept_delta']:+.4f} -> {'KEPT' if h['kept'] else 'rejected'}"
                         for h in self.history) or "(none)"
        msgs = [{"role": "user", "content": PROMPT.format(space=json.dumps(self.space), best=best,
                                                          config=json.dumps(best_cfg), history=hist)}]
        for _ in range(3):
            txt = self.researcher.samples(msgs, n=1, temperature=0.7, max_tokens=300)[0].text
            m = re.search(r"\{.*\}", txt or "", re.S)
            try:
                p = json.loads(m.group(0)) if m else None
            except json.JSONDecodeError:
                p = None
            if p and p.get("knob") in self.space and p.get("value") in self.space[p["knob"]] \
                    and best_cfg.get(p["knob"]) != p["value"] \
                    and not any(h["knob"] == p["knob"] and h["value"] == p["value"] for h in self.history):
                return p
        # fall back: first untried value of a random knob
        rng = np.random.default_rng(self.seed + len(self.history))
        for k in rng.permutation(list(self.space)):
            for v in self.space[k]:
                if best_cfg.get(k) != v and not any(h["knob"] == k and h["value"] == v for h in self.history):
                    return {"knob": k, "value": v, "hypothesis": "(fallback: untried value)"}
        return None

    def run(self, init_cfg: dict, qids: list) -> dict:
        S, A, H = self.split(qids)
        best_cfg = dict(init_cfg)
        best_s = self.evaluate(best_cfg, S + A)
        best_score = float(np.mean([best_s[q] for q in S]))
        for i in range(self.budget):
            p = self.propose(best_cfg, best_score)
            if p is None:
                break
            cfg = {**best_cfg, p["knob"]: p["value"]}
            s = self.evaluate(cfg, S + A)
            sd, se = self.paired(s, best_s, S)
            ad, _ = self.paired(s, best_s, A)
            kept = sd > self.z * se and ad >= 0
            self.history.append({"i": i, "knob": p["knob"], "value": p["value"], "hypothesis": p.get("hypothesis", ""),
                                 "search_delta": sd, "search_se": se, "accept_delta": ad, "kept": bool(kept),
                                 "search_score": float(np.mean([s[q] for q in S]))})
            if kept:
                best_cfg, best_s, best_score = cfg, s, float(np.mean([s[q] for q in S]))
        h0 = self.evaluate(dict(init_cfg), H)
        h1 = self.evaluate(best_cfg, H) if best_cfg != init_cfg else h0
        hd, hse = self.paired(h1, h0, H)
        return {"initial": init_cfg, "final": best_cfg, "n_search": len(S), "n_accept": len(A), "n_heldout": len(H),
                "heldout_initial": float(np.mean(list(h0.values()))), "heldout_final": float(np.mean(list(h1.values()))),
                "heldout_delta": hd, "heldout_se": hse, "attempts": len(self.history),
                "kept": sum(h["kept"] for h in self.history), "history": self.history}
