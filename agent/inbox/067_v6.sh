# TIMEOUT=300
# V6: InsightBench through pi with the three-model tools (queued behind the core jobs).
J=$(sbatch --parsable sbatch/V6_insightbench_pi.sbatch); echo "$J" > agent/state/V6.id; echo "V6: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
