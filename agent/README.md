# agent/ — the git-driven cluster runner (see docs/CLUSTER_HANDOFF.md section 3)

Branch: `claude/tabular-model-agent-exp-wvni5o` of YarinBeni/excel_magic, checked out at `~/projects/excel_magic` on the cluster.

- start (one line, per session): `sbatch sbatch/AGENT_runner.sbatch`
- stop: push an empty `agent/STOP`, or `scancel <jobid>`
- `agent/inbox/NNN_name.sh` runs once (bash -e, repo root, env active, `# TIMEOUT=` override), output to `agent/outbox/NNN_name.out`, script moved to `agent/done/`
- `agent/heartbeat`: pushed hourly when idle; older than ~2 h means the runner is gone
- `agent/state/*.id`: Slurm job ids written by inbox scripts so later scripts can chain dependencies
