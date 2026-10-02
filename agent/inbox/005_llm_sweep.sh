# TIMEOUT=300
# J5: the same search with 4 open LLMs from Hugging Face (Qwen3-Coder-30B-A3B, Qwen3-32B, GLM-4.5-Air, gpt-oss-20b), one GPU each, after J0.
J0=$(cat agent/state/J0.id 2>/dev/null || true)
[ -n "$J0" ] || { echo "no J0 id yet; run 001 first"; exit 1; }
if squeue -u "$USER" -h -o "%j" | grep -qE "^J5_llm_sweep$"; then echo "J5 already queued"; else
  J5=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J5_llm_sweep.sbatch); echo "$J5" > agent/state/J5.id; echo "J5 array: $J5 (after J0)"; fi
