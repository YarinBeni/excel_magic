# TIMEOUT=300
# J2 failed instantly (example rejected --harness heuristic; fixed). Cancel the remaining J2 tasks and resubmit.
J2=$(cat agent/state/J2.id 2>/dev/null || true); [ -n "$J2" ] && scancel "$J2" 2>/dev/null && echo "cancelled J2 $J2"
sleep 3
J2=$(sbatch --parsable sbatch/J2_tabarena.sbatch); echo "$J2" > agent/state/J2.id; echo "J2 array: $J2"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -14
