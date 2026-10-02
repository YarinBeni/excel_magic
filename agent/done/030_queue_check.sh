# TIMEOUT=120
# No result pushes for ~50 min: show the queue and the tail of each running job's log.
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
for f in $(ls -t sbatch/logs/J9_*.out sbatch/logs/J8_*_1.out 2>/dev/null | head -3); do echo "== $f"; grep -avE "^\s*$|it/s\]|%\|" "$f" | tail -8 | cut -c1-200; done
df -h "$HOME" | tail -1; git status --short | head -5; git log --oneline -1
