"""The two non-Claude harnesses, exercised without any LLM or model weights (logreg backbone)."""
import json
import types

from tabfm_auto.agent.heuristic import render_pipeline
from tabfm_auto.agent.search import run_search
from tabfm_auto.data import load_task
from tabfm_auto.pipeline import load_pipeline


def test_heuristic_template_is_a_valid_pipeline(tmp_path):
    p = tmp_path / "pipeline.py"
    p.write_text(render_pipeline({"regression": False, "count_encode": True, "crosses_top_k": 2, "prior_correction": True}))
    load_pipeline(p)  # validates the four hooks


def test_heuristic_search_end_to_end(tmp_path):
    t = load_task("iris")
    m = run_search(t, model_spec="logreg", harness="heuristic", budget_evals=4, budget_minutes=5, baselines=(),
                   name="h", run_dir=tmp_path / "run")
    assert m["n_evals"] >= 2 and m["best_cv"] <= m["p0_cv"] + 1e-12
    assert (tmp_path / "run" / "heuristic_trace.json").exists()


class _FakeClient:
    """Scripted OpenAI-compatible client: describe -> write -> eval -> finish."""

    def __init__(self):
        self.calls = 0
        self.chat = types.SimpleNamespace(completions=types.SimpleNamespace(create=self.create))

    def create(self, model, messages, tools, temperature):
        self.calls += 1
        script = [
            ("describe_data", {}),
            ("write_pipeline", {"source": open(messages[-1]["content"]).read() if False else _PIPE}),
            ("run_eval", {}),
            ("finish", {"notes": "kept identity"}),
        ]
        name, args = script[min(self.calls - 1, len(script) - 1)]
        tc = types.SimpleNamespace(id=f"c{self.calls}", function=types.SimpleNamespace(name=name, arguments=json.dumps(args)),
                                   model_dump=lambda: {"id": f"c{self.calls}", "type": "function",
                                                       "function": {"name": name, "arguments": json.dumps(args)}})
        msg = types.SimpleNamespace(content=None, tool_calls=[tc])
        return types.SimpleNamespace(choices=[types.SimpleNamespace(message=msg)],
                                     usage=types.SimpleNamespace(prompt_tokens=10, completion_tokens=5))


_PIPE = '''import numpy as np
MODEL_KWARGS = {}
def preprocess(X_train, y_train, X_test):
    return X_train, y_train, X_test
def engineer(X_train, y_train, X_test):
    X_train = X_train.copy(); X_test = X_test.copy()
    X_train["s"] = X_train.sum(axis=1); X_test["s"] = X_test.sum(axis=1)
    return X_train, X_test
def sample(X_train, y_train, X_test, max_rows):
    return [np.arange(len(X_train))]
def postprocess(pred, y_train_orig, X_test):
    return pred
'''


def test_openai_loop_with_fake_client(tmp_path):
    from tabfm_auto.agent.openai_compat import run_openai_agent
    from tabfm_auto.agent.search import build_workspace, make_split, tabfm_eval
    from tabfm_auto.logging_utils import RunLogger

    t = load_task("iris")
    with RunLogger("o", {}, run_dir=tmp_path / "run") as run:
        tr, _ = make_split(t, 0.3, 0)
        ws = build_workspace(run, t, tr, "logreg", 5, 5, 10000, 3, 0, 300)
        tabfm_eval(ws)  # P0
        info = run_openai_agent("task", ws, tmp_path / "run" / "agent.jsonl", model="fake", client=_FakeClient(),
                                max_turns=10, eval_timeout_s=300)
        run.finish({})
    assert info["tool_counts"] == {"describe_data": 1, "write_pipeline": 1, "run_eval": 1, "finish": 1}
    evals = [json.loads(line) for line in (ws / "evals.jsonl").read_text().splitlines()]
    assert len(evals) == 2 and evals[-1]["status"] == "ok" and evals[-1]["n_features"] == 5
    assert (ws / "NOTES.md").read_text() == "kept identity"


def test_cli_harness_with_fake_agent(tmp_path):
    """A 'coding agent' that is just a shell command: it writes a pipeline and calls tabfm-eval."""
    import shlex
    import sys

    t = load_task("iris")
    fake = tmp_path / "fake_agent.py"
    fake.write_text(
        "import pathlib, subprocess, sys\n"
        "pathlib.Path('pipeline.py').write_text(open(sys.argv[1]).read())\n"
        "subprocess.run([sys.executable, '-m', 'tabfm_auto.harness.cli'], check=False)\n"
    )
    pipe = tmp_path / "pipe.py"
    pipe.write_text(_PIPE)
    cmd = f"{shlex.quote(sys.executable)} {shlex.quote(str(fake))} {shlex.quote(str(pipe))} {{prompt_file}}"
    m = run_search(t, model_spec="logreg", harness="cli", agent_cmd=cmd, budget_evals=4, budget_minutes=5, baselines=(),
                   name="c", run_dir=tmp_path / "run")
    assert m["n_evals"] == 2 and m["agent"]["rc"] == 0
    assert (tmp_path / "run" / "agent_stream.log").read_text().startswith("# cmd:")
