# TIMEOUT=120
# P0 done (V1 51837). V5: SQL harness ablation A,B,C,D,F,SC + guarded auto-research, using V1's signals for the stack.
SIG="$HOME/projects/excel_magic/artifacts/runs/20261003T192850Z_V1_verifier_study_all/signals.csv"
ls -la "$SIG" || { echo "FAILED signals missing"; exit 1; }
J=$(sbatch --parsable --export=ALL,SIGNALS=$SIG sbatch/V5_sql_harness.sbatch); echo "$J" > agent/state/V5.id; echo "V5: $J"
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
