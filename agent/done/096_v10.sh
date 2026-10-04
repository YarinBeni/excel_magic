# TIMEOUT=60
sbatch sbatch/V10_gli_compare.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R"
