"""The VFD tools as a small local HTTP service, so any agent harness (pi, aider, Qwen Code, our loop) can call them from
a shell with `vfd <tool> key=value ...`. Models (GLiClass, the tabular backbones) stay loaded across calls.

  python -m vfd.toolserver --port 8765 --config F       (one process per GPU; serves many workspaces)
  vfd sql "SELECT category, count(*) FROM data GROUP BY 1"         (run inside a workspace with data.csv + meta.json)
  vfd deep_drivers target=category
  vfd text_labels column=short_description labels="printer;network outage;password reset"
  vfd record_insight text="Hardware is 67% of incidents" evidence_sql="SELECT ..."
  vfd ledger                                                     (the accepted insights so far)
"""
from __future__ import annotations

import argparse
import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pandas as pd

HELP = """Tools (run from this directory):
  vfd sql "<DuckDB SQL over table data>"
  vfd record_insight text="<one sentence with numbers>" evidence_sql="<SQL whose result shows it>"
  vfd ledger
{deep}{text}"""
HELP_DEEP = """  vfd deep_drivers target=<column> [exclude="a;b"]     which columns drive an outcome (frozen tabular model, all rows)
  vfd deep_what_if target=<column> column=<column>       how the predicted outcome changes with a column
  vfd deep_anomalies target=<column> [id_column=<col>]   rows most surprising given the rest
  vfd deep_drift time_column=<col> cutoff=<YYYY-MM-DD>   did the data change after the cutoff, and where
  vfd deep_predict target=<column> [time_column=<col>]   is there real signal for an outcome
"""
HELP_TEXT = """  vfd text_labels column=<text column> labels="label one;label two;..."   adds 0/1 columns <column>__<label>
"""


class State:
    def __init__(self, config: str, gliclass: str | None, deep_model: str, deep_fast: str):
        from .analyst import CONFIGS, Analyst

        self.cfg = CONFIGS[config]
        self.Analyst = Analyst
        self.gli = None
        if gliclass and (self.cfg.text or self.cfg.verify):
            from .signals import GLiClassScorer

            self.gli = GLiClassScorer(gliclass)
        self.deep = None
        if self.cfg.deep:
            from .deep import DeepTool

            self.deep = DeepTool(model=deep_model, fast=deep_fast, check="lightgbm", max_context=5000)
        self.sessions: dict[str, object] = {}
        self.lock = threading.Lock()   # one GPU: serialise tool calls

    def session(self, ws: str):
        import duckdb

        if ws not in self.sessions:
            df = pd.read_csv(Path(ws) / "data.csv")
            a = self.Analyst(chat=None, cfg=self.cfg, deep=self.deep, gli=self.gli)
            a.con = duckdb.connect()
            a.df = df
            a.con.register("data", a.df)
            a.ledger, a.rejected = [], []
            self.sessions[ws] = a
        return self.sessions[ws]

    def call(self, ws: str, tool: str, args: dict) -> str:
        with self.lock:
            a = self.session(ws)
            if tool == "ledger":
                return json.dumps({"accepted": [x["text"] for x in a.ledger], "rejected": len(a.rejected)})
            if tool == "help":
                return HELP.format(deep=HELP_DEEP if self.cfg.deep else "", text=HELP_TEXT if self.cfg.text else "")
            out = a.call(tool, args)
            a.log.append({"tool": tool, "args": args, "out": str(out)[:300]})
            return str(out)


def serve(port: int, state: State):
    class H(BaseHTTPRequestHandler):
        def do_POST(self):
            body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
            try:
                out = state.call(body["workspace"], body["tool"], body.get("args", {}))
            except Exception as e:
                out = f"ERROR {type(e).__name__}: {str(e)[:300]}"
            data = out.encode()
            self.send_response(200)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def log_message(self, *a):
            pass

    ThreadingHTTPServer(("127.0.0.1", port), H).serve_forever()


def cli(argv: list[str] | None = None) -> None:
    """`vfd <tool> [positional SQL] [key=value ...]`; lists as "a;b;c". Workspace = current directory."""
    import os
    import urllib.request

    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        argv = ["help"]
    tool, rest = argv[0], argv[1:]
    args: dict = {}
    for r in rest:
        if "=" in r and r.split("=", 1)[0].isidentifier():
            k, v = r.split("=", 1)
            args[k] = v.split(";") if k in ("labels", "exclude") else v
        elif tool == "sql":
            args["query"] = (args.get("query", "") + " " + r).strip()
    port = int(os.environ.get("VFD_PORT", "8765"))
    req = urllib.request.Request(f"http://127.0.0.1:{port}/", method="POST",
                                 data=json.dumps({"workspace": os.getcwd(), "tool": tool, "args": args}).encode())
    print(urllib.request.urlopen(req, timeout=900).read().decode())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8765)
    ap.add_argument("--config", default="F")
    ap.add_argument("--gliclass", default="knowledgator/gliclass-large-v3.0")
    ap.add_argument("--deep-model", default="kumo-tabular-l:n_estimators=2,device=cuda")
    ap.add_argument("--deep-fast", default="kumo-tabular-s:n_estimators=1,device=cuda")
    a = ap.parse_args()
    serve(a.port, State(a.config, a.gliclass, a.deep_model, a.deep_fast))


if __name__ == "__main__":
    main()
