"""OpenAI-compatible chat client for a local vLLM server: n samples with token logprobs, a yes/no judge read from
first-token logprobs, and a thread pool so a few thousand prompts finish in minutes."""
from __future__ import annotations

import math
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, Callable, Iterable


@dataclass
class Sample:
    text: str
    mean_logprob: float
    sum_logprob: float
    n_tokens: int


class Chat:
    def __init__(self, model: str, base_url: str | None = None, api_key: str | None = None, timeout: float = 300,
                 no_think: bool = False):
        from openai import OpenAI

        self.model = model
        self.client = OpenAI(base_url=base_url or os.environ.get("OPENAI_BASE_URL"),
                             api_key=api_key or os.environ.get("OPENAI_API_KEY", "EMPTY"), timeout=timeout)
        # GLM-4.5 / Qwen3 thinking models: switch thinking off for short structured answers
        self.extra = {"chat_template_kwargs": {"enable_thinking": False}} if no_think else {}

    def _call(self, **kw) -> Any:
        for attempt in range(4):
            try:
                return self.client.chat.completions.create(model=self.model, extra_body=self.extra or None, **kw)
            except Exception:
                if attempt == 3:
                    raise
                time.sleep(2 ** attempt)

    def samples(self, messages: list[dict], n: int = 1, temperature: float = 0.0, max_tokens: int = 1024) -> list[Sample]:
        r = self._call(messages=messages, n=n, temperature=temperature, max_tokens=max_tokens, logprobs=True)
        out = []
        for ch in r.choices:
            lps = [t.logprob for t in (ch.logprobs.content if ch.logprobs and ch.logprobs.content else [])]
            out.append(Sample(ch.message.content or "", float(sum(lps) / len(lps)) if lps else float("nan"),
                              float(sum(lps)) if lps else float("nan"), len(lps)))
        return out

    def yes_prob(self, messages: list[dict]) -> tuple[float, str]:
        """P(yes) from the first generated token's top logprobs (normalised over yes/no variants); falls back to the text."""
        r = self._call(messages=messages, n=1, temperature=0.0, max_tokens=4, logprobs=True, top_logprobs=10)
        ch = r.choices[0]
        text = (ch.message.content or "").strip()
        py = pn = 0.0
        if ch.logprobs and ch.logprobs.content:
            for t in ch.logprobs.content[0].top_logprobs:
                tok = t.token.strip().lower()
                if tok in ("yes", "y", "true"):
                    py += math.exp(t.logprob)
                elif tok in ("no", "n", "false"):
                    pn += math.exp(t.logprob)
        if py + pn > 0:
            return py / (py + pn), text
        # the model opened with something else (e.g. a reasoning block): let it finish and read its last yes / no
        r = self._call(messages=messages, n=1, temperature=0.0, max_tokens=600)
        full = re.sub(r"<think>.*?</think>", "", r.choices[0].message.content or "", flags=re.S).lower()
        words = re.findall(r"\b(yes|no)\b", full)
        return (1.0 if words and words[-1] == "yes" else 0.0 if words else 0.5), "fallback:" + full[-40:]


def pmap(fn: Callable, items: Iterable, workers: int = 32) -> list:
    items = list(items)
    with ThreadPoolExecutor(workers) as ex:
        return list(ex.map(fn, items))


_SQL_BLOCK = re.compile(r"```(?:sql|sqlite)?\s*(.*?)```", re.S | re.I)


def extract_sql(text: str) -> str:
    """Last fenced SQL block, else the text from the first SELECT/WITH; trailing semicolons removed."""
    blocks = _SQL_BLOCK.findall(text or "")
    sql = blocks[-1] if blocks else text or ""
    if not blocks:
        m = re.search(r"\b(WITH|SELECT)\b", sql, re.I)
        sql = sql[m.start():] if m else sql
    return sql.strip().rstrip(";").strip()
