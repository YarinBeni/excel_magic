# TIMEOUT=120
# Smokes passed (V1 30 questions, V4 rel-f1 driver-dnf). Full runs: V1 on all 498 Arcwise-Plat questions; V4 on 8 tasks x 2 repeats.
J1=$(sbatch --parsable sbatch/V1_verifier_study.sbatch); echo "$J1" > agent/state/V1.id; echo "V1 full: $J1"
J4=$(sbatch --parsable sbatch/V4_relbench_capability.sbatch); echo "$J4" > agent/state/V4.id; echo "V4 full: $J4"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
