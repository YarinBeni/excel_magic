# Cluster work branch: open-tabfm-auto experiments

This branch of `excel_magic` is used ONLY as the cluster's command channel and lab notebook for the TabFM-Auto
experiments (the repos will move to their own GitHub homes later). `main` still holds the Excel duplicate finder
(moved here to `legacy_excel_magic/`).

- `open-tabfm-auto/`: the library (open-source TabFM-Auto)
- `frozen-embeddings-retrieval/`: the embeddings research repo
- `sbatch/`: Slurm jobs J0 (setup) ... J4 (RelBench probe), `_common.sh` prologue, `_sandbox.sh` git helpers
- `agent/`: git-driven runner (inbox / outbox / done / heartbeat)
- `reports/`: results pushed back by jobs (`reports/runs/<run>/metrics.json`, `reports/logs/`)
- `artifacts/` (gitignored): weights, OpenML cache, full run dirs, vLLM logs
- `STATUS.md`: living status board; `docs/CLUSTER_HANDOFF.md`: the operating manual

## One-time on the cluster (Yarin)
```
git clone -b claude/tabular-model-agent-exp-wvni5o https://github.com/YarinBeni/excel_magic ~/projects/excel_magic && cd ~/projects/excel_magic && git remote set-url origin https://<TOKEN>@github.com/YarinBeni/excel_magic && mkdir -p sbatch/logs && sbatch sbatch/AGENT_runner.sbatch && squeue -u $USER | head
```
## Per session
```
cd ~/projects/excel_magic && sbatch sbatch/AGENT_runner.sbatch
```
