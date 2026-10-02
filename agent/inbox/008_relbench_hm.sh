# TIMEOUT=300
# J7: RelBench rel-hm user-item-purchase with our frozen-embedding kNN rows (official MAP@12). Needs only J0 (done).
if squeue -u "$USER" -h -o "%j" | grep -qE "^J7_relbench_hm$"; then echo "J7 already queued"; else
  J7=$(sbatch --parsable sbatch/J7_relbench_hm.sbatch); echo "$J7" > agent/state/J7.id; echo "J7: $J7"; fi
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
