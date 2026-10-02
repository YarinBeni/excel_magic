# TIMEOUT=300
# J5 task 3 (gpt-oss-20b, 50376): openai_harmony could not write /tmp/tiktoken-rs-cache (owned by another user on the node).
# vllm_env now points TIKTOKEN_RS_CACHE_DIR at ~/.cache. Resubmit array task 3 only.
J5c=$(sbatch --parsable --array=3 sbatch/J5_llm_sweep.sbatch); echo "$J5c" > agent/state/J5_task3.id; echo "J5 task3: $J5c"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
