# TIMEOUT=60
sbatch sbatch/V9_gliclass_ft.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %.12l %R" | head -4
