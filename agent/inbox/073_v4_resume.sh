# TIMEOUT=120
scancel 51838; sleep 5
RUN=$(ls -td artifacts/runs/*V4_relbench_capability | head -1); echo "resume $RUN"; wc -l < "$RUN/capability_rows.jsonl"
RUN="$RUN" TASKS=rel-avito/user-clicks,rel-hm/user-churn sbatch sbatch/V4_relbench_capability.sbatch
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
