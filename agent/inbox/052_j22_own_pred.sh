# TIMEOUT=120
# 051 was killed (exit 137, out of memory on the runner node) building rel-event; run the check as a GPU job instead.
J22=$(sbatch --parsable sbatch/J22_kumo_own_pred_check.sbatch); echo "$J22" > agent/state/J22.id; echo "J22: $J22"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
