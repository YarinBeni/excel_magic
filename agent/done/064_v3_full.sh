# TIMEOUT=120
# V3 smoke passed (5 tables). Same insight-count instruction for all configs (F recorded 2.4 vs D 6.4 with 0 rejections),
# then the full 100-table run.
J=$(sbatch --parsable sbatch/V3_insightbench.sbatch); echo "$J" > agent/state/V3.id; echo "V3 full: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
