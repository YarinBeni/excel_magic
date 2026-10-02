# TIMEOUT=300
# J6 task 0 (pi, 50377_0) failed in 4 min: `conda create -n node22` raced with task 1 creating the same env (corrupted-package
# errors). The env exists now (Qwen Code ran on it) and creation is under a lock. Resubmit array task 0 only.
J6p=$(sbatch --parsable --array=0 sbatch/J6_cli_agents.sbatch); echo "$J6p" > agent/state/J6_task0.id; echo "J6 task0 (pi): $J6p"
squeue -u "$USER" -o "%.10F %.14j %.8T %.10M %R" | sort -u | head -8
