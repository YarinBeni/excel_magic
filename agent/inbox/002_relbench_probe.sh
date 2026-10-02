# TIMEOUT=300
# J4: RelBench rel-hm probe (CPU), after J0. Independent of J1.
J0=$(cat agent/state/J0.id 2>/dev/null || true)
[ -n "$J0" ] || { echo "no J0 id yet; run 001 first"; exit 1; }
if squeue -u "$USER" -h -o "%j" | grep -qE "^J4_relbench$"; then echo "J4 already queued"; else
  J4=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J4_relbench_probe.sbatch); echo "$J4" > agent/state/J4.id; echo "J4: $J4 (after J0)"; fi
