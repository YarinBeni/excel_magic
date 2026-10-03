# TIMEOUT=120
J=$(sbatch --parsable --export=ALL,TASKS=rel-f1/driver-dnf,REPEATS=1 --time=04:00:00 sbatch/V4_relbench_capability.sbatch); echo "$J" > agent/state/V4smoke.id; echo "V4 smoke: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
