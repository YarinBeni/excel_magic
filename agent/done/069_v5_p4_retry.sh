# TIMEOUT=120
# V5 P4 crashed: the researcher used the generator's model name on the judge server. Fixed; rerun P4 only.
SIG="$HOME/projects/excel_magic/artifacts/runs/20261003T192850Z_V1_verifier_study_all/signals.csv"
J=$(sbatch --parsable --export=ALL,SIGNALS=$SIG,SKIP_P2=1 sbatch/V5_sql_harness.sbatch); echo "$J" > agent/state/V5p4.id; echo "V5 P4: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
