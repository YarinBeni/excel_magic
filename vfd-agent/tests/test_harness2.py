import sqlite3
from types import SimpleNamespace

from vfd import bird
from vfd.harness2 import CONFIGS2, answer2, literal_issues, profile_schema


def _db(tmp_path):
    p = tmp_path / "shop" / "shop.sqlite"
    p.parent.mkdir()
    con = sqlite3.connect(p)
    con.execute("CREATE TABLE customers (id INTEGER PRIMARY KEY, segment TEXT, age INTEGER)")
    con.executemany("INSERT INTO customers VALUES (?,?,?)", [(i, ["SME", "LAM", "KAM"][i % 3], 20 + i) for i in range(60)])
    con.commit()
    con.close()
    return p


class FakeChat:
    def __init__(self, replies):
        self.replies, self.calls = list(replies), []

    def samples(self, msgs, n=1, temperature=0.0, max_tokens=1024):
        self.calls.append(msgs)
        return [SimpleNamespace(text=self.replies.pop(0)) for _ in range(n)]


def test_profile_and_literals(tmp_path):
    p = _db(tmp_path)
    s = profile_schema(p, bird.load_schema(p))
    txt = bird.render_schema(s)
    assert "3 distinct" in txt and "range 20 to 79" in txt
    assert literal_issues(p, "SELECT * FROM customers WHERE segment = 'SME'") == []
    msg = literal_issues(p, "SELECT * FROM customers WHERE segment = 'sme'")
    assert msg and "'SME'" in msg[0]
    assert literal_issues(p, "SELECT * FROM customers WHERE segment LIKE 'S%'") == []


def test_gates_revise(tmp_path):
    p = _db(tmp_path)
    q = bird.Question(1, "shop", "How many SME customers?", "", "SELECT COUNT(*) FROM customers WHERE segment='SME'")
    chat = FakeChat(["```sql\nSELECT COUNT(*) FROM customers WHERE segment = 'sme'\n```",
                     "```sql\nSELECT COUNT(*) FROM customers WHERE segment = 'SME'\n```"])
    r = answer2(q, p, bird.render_schema(bird.load_schema(p)), CONFIGS2["G"], chat)
    assert r["revisions"] == 1 and r["passed_gates"]
    assert r["result_key"] == bird.execute(p, q.gold_sql).key
    assert "not a stored value" in chat.calls[1][-1]["content"]
    # the rule card is in the system prompt only for configs with rules
    assert "Rules:" in chat.calls[0][0]["content"]
    chat = FakeChat(["```sql\nSELECT 1\n```"])
    answer2(q, p, "", CONFIGS2["A"], chat)
    assert "Rules:" not in chat.calls[0][0]["content"]
