# TIMEOUT=300
# J19: TabPFN v2 layer sweep on the same RelBench tasks as J18. J20: TEmBed row benchmark with TabPFN context-target variants.
for j in J19_relbench_tab_layers J20_tembed_rows; do
  if squeue -u "$USER" -h -o "%j" | grep -q "^$j\$"; then echo "$j already queued"; continue; fi
  id=$(sbatch --parsable sbatch/$j.sbatch); echo "$id" > "agent/state/${j%%_*}.id"; echo "$j: $id"
done
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
