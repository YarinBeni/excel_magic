# TIMEOUT=300
# J7c: rerun rel-hm with the new reference rows (user-kNN on the raw purchase matrix, item-kNN, past-only+fill hybrids)
# so the negative result separates "weak embeddings" from "weak kNN scoring". Idempotent: skip if a J7 is still queued.
if squeue -u "$USER" -h -o "%j" | grep -q '^J7_relbench_hm$'; then echo "J7 already queued/running"; exit 0; fi
J7=$(sbatch --parsable sbatch/J7_relbench_hm.sbatch); echo "$J7" > agent/state/J7.id; echo "J7c: $J7"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -12
