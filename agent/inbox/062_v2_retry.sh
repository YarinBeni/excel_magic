# TIMEOUT=120
J=$(sbatch --parsable sbatch/V2_text_featurizer.sbatch); echo "$J" > agent/state/V2.id; echo "V2: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
