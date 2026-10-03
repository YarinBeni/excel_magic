# TIMEOUT=300
# J21: all six open tabular FMs, every layer, 5 RelBench entity tasks: probes, label-free geometry, CKA.
if squeue -u "$USER" -h -o "%j" | grep -q '^J21_model_layers$'; then echo "J21 already queued"; exit 0; fi
J21=$(sbatch --parsable sbatch/J21_model_layers.sbatch); echo "$J21" > agent/state/J21.id; echo "J21: $J21"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
