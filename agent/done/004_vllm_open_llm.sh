# TIMEOUT=300
# J3: the search driven by an open LLM (Qwen3-Coder-30B-A3B via vLLM) on one GPU, after J0. Long (install + download + runs).
J0=$(cat agent/state/J0.id 2>/dev/null || true)
[ -n "$J0" ] || { echo "no J0 id yet; run 001 first"; exit 1; }
if squeue -u "$USER" -h -o "%j" | grep -qE "^J3_vllm$"; then echo "J3 already queued"; else
  J3=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J3_vllm_search.sbatch); echo "$J3" > agent/state/J3.id; echo "J3: $J3 (after J0)"; fi
