# TIMEOUT=300
# Next-steps batch: J9 per-backbone TabArena (Kumo-L, TabICLv2), J10 budget 64 on the paper's high-gain datasets (3 setups),
# J11 repeated-CV judge on small tables (heuristic, pi), J12 rel-hm user-churn probe, J13 rich tool loop. 9 GPU tasks.
for j in J9_tabarena_backbones J10_budget J11_judge J12_relbench_churn J13_loop_rich; do
  if squeue -u "$USER" -h -o "%j" | grep -q "^$j\$"; then echo "$j already queued"; continue; fi
  id=$(sbatch --parsable sbatch/$j.sbatch); echo "$id" > "agent/state/${j%%_*}.id"; echo "$j: $id"
done
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u | head -20
