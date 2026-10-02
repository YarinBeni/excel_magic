# TIMEOUT=300
# J8 task 1 (GLM-4.5-Air-FP8): 128k context needs 23 GiB KV, only 18 GiB left next to the FP8 weights -> 64k. Resubmit task 1.
J8g=$(sbatch --parsable --array=1 sbatch/J8_tabarena_llm.sbatch); echo "$J8g" > agent/state/J8_task1.id; echo "J8 task1 (GLM): $J8g"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -8
