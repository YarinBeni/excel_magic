# TIMEOUT=60
CONFIGS=D,Q0,Q sbatch sbatch/V3_insightbench.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
