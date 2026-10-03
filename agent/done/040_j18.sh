# TIMEOUT=300
# J18: layer-wise entity embeddings of frozen Kumo Relational on 8 real RelBench entity-classification tasks.
if squeue -u "$USER" -h -o "%j" | grep -q '^J18_relbench_layers$'; then echo "J18 already queued"; exit 0; fi
J18=$(sbatch --parsable sbatch/J18_relbench_layers.sbatch); echo "$J18" > agent/state/J18.id; echo "J18: $J18"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
