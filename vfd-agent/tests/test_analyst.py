import json
from types import SimpleNamespace as NS

import numpy as np
import pandas as pd

from vfd.analyst import CONFIGS, Analyst, g_eval_match


class FakeClient:
    """Replays a fixed list of tool calls, one per completion."""
    def __init__(self, script):
        self.script = list(script)
        self.chat = NS(completions=NS(create=self.create))

    def create(self, **kw):
        name, args = self.script.pop(0) if self.script else ("finish", {"summary": "done"})
        tc = NS(id=f"c{len(self.script)}", type="function", function=NS(name=name, arguments=json.dumps(args)))
        tc.model_dump = lambda: {"id": tc.id, "type": "function", "function": {"name": name, "arguments": json.dumps(args)}}
        return NS(choices=[NS(message=NS(content="", tool_calls=[tc]))])


class FakeChat:
    def __init__(self, script):
        self.client = FakeClient(script)
        self.model = "fake"


class FakeGli:
    def scores(self, texts, labels):
        return np.array([[0.9 if ("printer" in t.lower() and "printer" in lab.lower()) or lab.lower().startswith("hardware")
                          else 0.05 for lab in labels] for t in texts], np.float32)


def table():
    return pd.DataFrame({"category": ["Hardware"] * 6 + ["Software"] * 2,
                         "desc": ["printer jam"] * 5 + ["screen"] + ["crash"] * 2,
                         "days": [1, 2, 3, 4, 5, 6, 7, 8]})


def test_unverified_config_records_anything():
    script = [("sql", {"query": "SELECT category, count(*) n FROM data GROUP BY 1"}),
              ("record_insight", {"text": "Hardware has 99 incidents", "evidence_sql": "SELECT 1"})]
    r = Analyst(FakeChat(script), CONFIGS["D"]).run(table(), {"goal": "g"})
    assert r["insights"] == ["Hardware has 99 incidents"] and r["n_rejected"] == 0


def test_verified_ledger_rejects_wrong_numbers_and_accepts_right_ones():
    ev = "SELECT category, count(*) n FROM data GROUP BY 1 ORDER BY 2 DESC"
    script = [("record_insight", {"text": "Hardware has 99 incidents", "evidence_sql": ev}),
              ("record_insight", {"text": "Hardware has 6 incidents, 75% of all", "evidence_sql":
                                  "SELECT count(*) FILTER (WHERE category='Hardware') n, "
                                  "count(*) FILTER (WHERE category='Hardware')::DOUBLE / count(*) AS frac FROM data"}),
              ("text_labels", {"column": "desc", "labels": ["printer problem", "software crash"]}),
              ("sql", {"query": "SELECT sum(desc__printer_problem) FROM data"})]
    a = Analyst(FakeChat(script), CONFIGS["F"], gli=FakeGli())
    r = a.run(table(), {"goal": "g"})
    assert r["n_rejected"] == 1 and r["insights"] == ["Hardware has 6 incidents, 75% of all"]
    assert "desc__printer_problem" in a.df.columns and a.df["desc__printer_problem"].sum() == 5
    assert "5" in a.log[-2]["out"] or "5" in a.log[-1]["out"]


def test_g_eval_match():
    class J:
        def samples(self, msgs, n=1, temperature=0.0, max_tokens=8):
            return [NS(text="10" if "printer" in msgs[0]["content"].split("Predicted")[0] else "1")]
    s = g_eval_match(J(), ["printers dominate hardware"], ["most hardware incidents are printer issues", "trend up"])
    assert s["per_gold"] == [1.0, 0.0] and s["g_eval"] == 0.5


def test_number_supported():
    from vfd.analyst import number_supported
    assert number_supported("6", [6.0, 2.0]) and not number_supported("99", [6.0, 2.0, 1.0])
    assert number_supported("75%", [0.75]) and number_supported("33.3%", [0.33333]) and not number_supported("40%", [0.75])
    assert number_supported("12.5", [12.46]) and not number_supported("12.5", [12.4])
    assert not number_supported("75", [0.75])  # a bare number does not match a fraction
