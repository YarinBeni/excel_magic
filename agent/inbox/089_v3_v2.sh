# TIMEOUT=60
CONFIGS=D,P,PC sbatch sbatch/V3_insightbench.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
