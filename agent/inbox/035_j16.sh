# TIMEOUT=300
# J16: the identity pipeline scores differently in J15 than in J2/J8 on the same splits -> determinism / library-version check.
J16=$(sbatch --parsable sbatch/J16_determinism.sbatch); echo "$J16" > agent/state/J16.id; echo "J16: $J16"
