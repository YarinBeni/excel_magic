# TIMEOUT=300
# J1: rerun with fixed model-list parsing (Kumo-M/L, EXAONE on cuda), transfer discovery, float32 GPU embeddings,
#     Kumo Relational with k-means target. J7: rerun with subset-safe evaluation (smoke) and test-split evaluation.
J1=$(sbatch --parsable sbatch/J1_smoke.sbatch); echo "$J1" > agent/state/J1.id; echo "J1b: $J1"
J7=$(sbatch --parsable sbatch/J7_relbench_hm.sbatch); echo "$J7" > agent/state/J7.id; echo "J7b: $J7"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -16
