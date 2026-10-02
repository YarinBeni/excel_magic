# TIMEOUT=300
# J6: free coding-agent CLIs (pi, Qwen Code, aider) driving Qwen3-Coder-30B-A3B via vLLM, one GPU each, after J0.
J0=$(cat agent/state/J0.id 2>/dev/null || true)
[ -n "$J0" ] || { echo "no J0 id yet; run 001 first"; exit 1; }
if squeue -u "$USER" -h -o "%j" | grep -qE "^J6_cli_agents$"; then echo "J6 already queued"; else
  J6=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J6_cli_agents.sbatch); echo "$J6" > agent/state/J6.id; echo "J6 array: $J6 (after J0)"; fi
