# TIMEOUT=120
# V3 full: the insight checker rejected correct derived percentages (6 of 372 = 1.6%). Fixed; rerun config F only.
J=$(sbatch --parsable --export=ALL,CONFIGS=F sbatch/V3_insightbench.sbatch); echo "$J" > agent/state/V3F.id; echo "V3 F rerun: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
