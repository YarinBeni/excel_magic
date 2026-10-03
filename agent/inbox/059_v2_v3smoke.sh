# TIMEOUT=120
# P1 featurizer study (full, 18 datasets) and a 5-table smoke of the InsightBench analyst (configs D, E, F).
J2=$(sbatch --parsable sbatch/V2_text_featurizer.sbatch); echo "$J2" > agent/state/V2.id; echo "V2: $J2"
J3=$(sbatch --parsable --export=ALL,LIMIT=5 --time=04:00:00 sbatch/V3_insightbench.sbatch); echo "$J3" > agent/state/V3smoke.id; echo "V3 smoke: $J3"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
