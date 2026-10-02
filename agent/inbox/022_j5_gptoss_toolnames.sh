# TIMEOUT=300
# J5 task 3 (gpt-oss-20b, 50431) ran, but vLLM's harmony parser mangled tool names and the TabArena run hit choices=None.
# Harness normalises names and guards empty responses; resubmit array task 3 for a clean run.
J5e=$(sbatch --parsable --array=3 sbatch/J5_llm_sweep.sbatch); echo "$J5e" > agent/state/J5_task3.id; echo "J5 task3: $J5e"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
