# TIMEOUT=300
# J0 (env + weights + tests) then J1 (first GPU results) after J0 succeeds. Idempotent: skips if a J0/J1 is already queued.
mkdir -p agent/state sbatch/logs
if squeue -u "$USER" -h -o "%j" | grep -qE "^J0_setup$"; then echo "J0 already queued/running: $(cat agent/state/J0.id 2>/dev/null)"; else
  J0=$(sbatch --parsable sbatch/J0_setup.sbatch); echo "$J0" > agent/state/J0.id; echo "J0: $J0"; fi
J0=$(cat agent/state/J0.id)
if squeue -u "$USER" -h -o "%j" | grep -qE "^J1_smoke$"; then echo "J1 already queued"; else
  J1=$(sbatch --parsable --dependency=afterok:"$J0" sbatch/J1_smoke.sbatch); echo "$J1" > agent/state/J1.id; echo "J1: $J1 (after J0)"; fi
squeue -u "$USER" -o "%.10i %.14j %.8T %.10M %R" | head -10
