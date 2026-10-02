# TIMEOUT=300
# J12c: the supervised-TabPFN row hit the TabPFN license prompt (direct TabPFNClassifier); it now goes through the model
# registry like the embedders. Rerun the probe for that one reference row.
J12c=$(sbatch --parsable sbatch/J12_relbench_churn.sbatch); echo "$J12c" > agent/state/J12.id; echo "J12c: $J12c"
