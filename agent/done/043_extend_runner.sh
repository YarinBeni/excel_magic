# TIMEOUT=120
# Yarin asked (2026-10-03) to keep the runner alive 2 more days. Try to raise the running job's limit; users usually may
# only lower it, so fall back to a 2-day successor that starts when this runner ends (afterany): no gap, never two runners.
cur=$(cat agent/state/RUNNER.id 2>/dev/null || squeue -u "$USER" -h -n AGENT_runner -o "%i" | head -1)
echo "current runner: $cur"; sinfo -p voltagepark -h -o "partition max time: %l"
squeue -j "$cur" -h -o "time used %M of %l, ends %e"
if scontrol update JobId="$cur" TimeLimit=3-00:00:00 2>&1; then
    echo "raised in place:"; squeue -j "$cur" -h -o "time used %M of %l, ends %e"
else
    if squeue -u "$USER" -h -n AGENT_runner -o "%i %T %r" | grep -q "Dependency"; then echo "successor already queued"; else
        nxt=$(sbatch --parsable --dependency=afterany:"$cur" sbatch/AGENT_runner_2d.sbatch); echo "$nxt" > agent/state/RUNNER_next.id
        echo "successor runner $nxt (2 days) queued after $cur"
    fi
fi
squeue -u "$USER" -n AGENT_runner -o "%.10i %.14j %.10T %.10M %.12l %R"
