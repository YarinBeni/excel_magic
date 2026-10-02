# TIMEOUT=600
# J6 task 1 (qwen-code, 50408): the qwen CLI exited 1 after 2.3 s on every dataset and its log is not in the pushed files.
# Print that log now (the artifacts dir is still on disk), then resubmit task 1 with the agent log kept + echoed.
for d in artifacts/runs/*J6_qwen-code_synth_physics; do echo "== $d/agent_stream.log"; head -60 "$d/agent_stream.log" 2>/dev/null; done
export PATH="$HOME/miniconda3/envs/node22/bin:$PATH"; command -v qwen && qwen --version 2>&1 | head -3; qwen --help 2>&1 | grep -E "max-wall|max-session|output-format|approval" | head -6
J6b=$(sbatch --parsable --array=1 sbatch/J6_cli_agents.sbatch); echo "$J6b" > agent/state/J6_task1.id; echo "J6 task1: $J6b"
