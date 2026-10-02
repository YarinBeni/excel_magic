# TIMEOUT=300
# J12b: the supervised-TabPFN reference row failed (inference_precision wants a torch dtype). Fixed; rerun the churn probe.
J12b=$(sbatch --parsable sbatch/J12_relbench_churn.sbatch); echo "$J12b" > agent/state/J12.id; echo "J12b: $J12b"
