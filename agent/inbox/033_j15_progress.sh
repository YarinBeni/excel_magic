# TIMEOUT=120
# J15 has pushed nothing in ~3 h: show its progress lines and the queue.
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
f=$(ls -t sbatch/logs/J15_*.out | head -1); echo "== $f"; grep -a "\[rescore\]\|FAILED\|Error\|Traceback" "$f" | tail -12 | cut -c1-220; tail -3 "$f" | cut -c1-220
f=$(ls -t sbatch/logs/J14b_*.out 2>/dev/null | head -1); [ -n "$f" ] && { echo "== $f"; tail -3 "$f" | cut -c1-200; }
