# TIMEOUT=120
squeue -u "$USER" -o "%.12i %.22j %.8T %.10M %R" | sort -u
d=$(ls -td artifacts/runs/*V4_relbench_capability | head -1); echo "== V4 $d"; wc -l < "$d/capability_rows.jsonl" 2>/dev/null; tail -3 "$d/capability_rows.jsonl" 2>/dev/null | cut -c1-200
f=$(ls -t sbatch/logs/V4_rel_capability_51838*.out | head -1); grep -a "^\[cap\]" "$f" | tail -4 | cut -c1-200
d=$(ls -td artifacts/runs/*V6_insightbench_pi_all | head -1); echo "== V6 $d"; for c in D E F; do printf "%s " $c; wc -l < "$d/insights_pi$c.jsonl" 2>/dev/null || echo 0; done
f=$(ls -t sbatch/logs/V6_insight_pi_51914*.out | head -1); grep -a "^\[pi\]" "$f" | tail -3 | cut -c1-200
