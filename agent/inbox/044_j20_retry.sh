# TIMEOUT=300
# J20 failed: the TEmBed repo contains a folder named "<approach_name>" that this filesystem rejects. Sparse checkout now
# skips it (and removes the half clone). Resubmit J20.
J20=$(sbatch --parsable sbatch/J20_tembed_rows.sbatch); echo "$J20" > agent/state/J20.id; echo "J20: $J20"
