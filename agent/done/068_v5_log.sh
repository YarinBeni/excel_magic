# TIMEOUT=120
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
f=$(ls -t sbatch/logs/V5_sql_harness_51863*.out 2>/dev/null | head -1); echo "== $f"
grep -a -A25 "P4 auto-research" "$f" | grep -av "Fetching\|it/s\]" | cut -c1-300 | tail -30
