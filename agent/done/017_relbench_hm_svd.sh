# TIMEOUT=300
# J7d: rel-hm with the interaction signal as the TFM input (SVD of the purchase matrix) vs the raw factors. One GPU, ~1h.
if squeue -u "$USER" -h -o "%j" | grep -q '^J7d_relbench_hm_svd$'; then echo "J7d already queued/running"; exit 0; fi
J7d=$(sbatch --parsable sbatch/J7d_relbench_hm_svd.sbatch); echo "$J7d" > agent/state/J7d.id; echo "J7d: $J7d"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -12
