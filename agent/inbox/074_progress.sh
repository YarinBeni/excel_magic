# TIMEOUT=120
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R"
f=$(ls -t sbatch/logs/V4_rel_capability_52039*.out | head -1); echo "== $f"; grep -a "^\[cap\]\|FAILED\|Error\|Traceback" "$f" | tail -6 | cut -c1-220
d=$(ls -td artifacts/runs/*V6_insightbench_pi_all | head -1); for c in D E F; do printf "%s " $c; wc -l < "$d/insights_pi$c.jsonl"; done
