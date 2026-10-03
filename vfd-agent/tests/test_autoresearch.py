import json
from types import SimpleNamespace as NS

import numpy as np

from vfd.autoresearch import AutoResearch


class Researcher:
    def __init__(self, plan):
        self.plan = list(plan)

    def samples(self, msgs, n=1, temperature=0.7, max_tokens=300):
        p = self.plan.pop(0) if self.plan else {"knob": "x", "value": 0}
        return [NS(text=json.dumps(p))]


def make_eval(effects, noise=0.0, seed=0):
    """score of question q under cfg: base 0.5 + sum of knob effects, as a Bernoulli draw fixed per (q, cfg)."""
    def ev(cfg, qids):
        out = {}
        for q in qids:
            p = 0.5 + sum(effects.get((k, v), 0.0) for k, v in cfg.items())
            r = np.random.default_rng(hash((q, json.dumps(cfg, sort_keys=True))) % 2**32).random()
            out[q] = float(r < p)
        return out
    return ev


def test_keeps_a_real_gain_and_rejects_noise():
    space = {"n": [1, 8], "k": [10, 20, 40]}
    ev = make_eval({("n", 8): 0.35})               # a large real effect; k does nothing
    ar = AutoResearch(ev, space, Researcher([{"knob": "k", "value": 40}, {"knob": "n", "value": 8},
                                              {"knob": "k", "value": 20}]), budget=3, z=1.0)
    res = ar.run({"n": 1, "k": 10}, list(range(600)))
    assert res["final"]["n"] == 8
    assert res["heldout_final"] > res["heldout_initial"]
    assert [h["kept"] for h in res["history"]].count(True) == 1


def test_no_repeats_and_fallback():
    ar = AutoResearch(make_eval({}), {"a": [0, 1]}, Researcher([{"knob": "a", "value": 1}, {"knob": "a", "value": 1}]),
                      budget=3)
    res = ar.run({"a": 0}, list(range(100)))
    assert res["attempts"] <= 2
