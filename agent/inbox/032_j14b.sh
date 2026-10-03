# TIMEOUT=300
# J14b: item-level rows for Kumo Relational with the random in-context target (the churn probe preferred random over k-means).
J14b=$(sbatch --parsable sbatch/J14b_relational_item_random.sbatch); echo "$J14b" > agent/state/J14b.id; echo "J14b: $J14b"
