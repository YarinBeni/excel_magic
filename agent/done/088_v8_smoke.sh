# TIMEOUT=60
TASKS=rel-f1/driver-dnf REPEATS=1 BUDGET=6 sbatch sbatch/V8_relbench_hypo.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
