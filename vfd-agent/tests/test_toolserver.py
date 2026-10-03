import json
import threading
import time

import numpy as np
import pandas as pd

from vfd import toolserver


class FakeGli:
    def scores(self, texts, labels):
        return np.array([[0.9 if ("printer" in t.lower() and "printer" in lab.lower()) or lab.startswith("Hardware") else 0.05 for lab in labels]
                         for t in texts], np.float32)


def test_server_and_cli(tmp_path, monkeypatch, capsys):
    ws = tmp_path / "ws"
    ws.mkdir()
    pd.DataFrame({"category": ["Hardware"] * 3 + ["Software"], "desc": ["printer jam"] * 3 + ["crash"]}).to_csv(ws / "data.csv", index=False)
    st = toolserver.State.__new__(toolserver.State)
    from vfd.analyst import CONFIGS, Analyst
    st.cfg, st.Analyst, st.gli, st.deep, st.sessions, st.lock = CONFIGS["F"], Analyst, FakeGli(), None, {}, threading.Lock()
    port = 8799
    threading.Thread(target=toolserver.serve, args=(port, st), daemon=True).start()
    import socket
    for _ in range(100):  # wait until the server accepts connections
        try:
            socket.create_connection(("127.0.0.1", port), timeout=0.1).close()
            break
        except OSError:
            time.sleep(0.05)
    monkeypatch.setenv("VFD_PORT", str(port))
    monkeypatch.chdir(ws)
    toolserver.cli(["sql", "SELECT", "category,", "count(*)", "n", "FROM", "data", "GROUP", "BY", "1", "ORDER", "BY", "2", "DESC"])
    assert "Hardware" in capsys.readouterr().out
    toolserver.cli(["text_labels", "column=desc", "labels=printer problem;crash"])
    assert "desc__printer_problem" in capsys.readouterr().out
    toolserver.cli(["record_insight", "text=Hardware has 3 incidents",
                    "evidence_sql=SELECT count(*) AS n FROM data WHERE category='Hardware'"])
    assert "ACCEPTED" in capsys.readouterr().out
    toolserver.cli(["ledger"])
    assert json.loads(capsys.readouterr().out)["accepted"] == ["Hardware has 3 incidents"]
