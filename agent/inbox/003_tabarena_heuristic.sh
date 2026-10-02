# TIMEOUT=300
# J2: TabArena protocol, 17 small datasets, heuristic harness, Kumo Tabular-S on GPU. Array of 17 one-GPU tasks, after J0.
J0=$(cat agent/state/J0.id 2>/dev/null || true)
[ -n "$J0" ] || { echo "no J0 id yet; run 001 first"; exit 1; }
if squeue -u "$USER" -h -o "%j" | grep -qE "^J2_tabarena$"; then echo "J2 already queued"; else
  J2=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J2_tabarena.sbatch); echo "$J2" > agent/state/J2.id; echo "J2 array: $J2 (after J0)"; fi
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -12
