# TIMEOUT=300
# J14: Kumo Relational on rel-hm (churn probe + item-level rows). J15: conservative selection rules re-scored on the
# finished TabArena searches (no new agent runs).
for j in J14_relational_relbench J15_selection_rules; do
  if squeue -u "$USER" -h -o "%j" | grep -q "^$j\$"; then echo "$j already queued"; continue; fi
  id=$(sbatch --parsable sbatch/$j.sbatch); echo "$id" > "agent/state/${j%%_*}.id"; echo "$j: $id"
done
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u | head -12
