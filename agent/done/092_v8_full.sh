# TIMEOUT=60
REPEATS=2 BUDGET=10 sbatch sbatch/V8_relbench_hypo.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
