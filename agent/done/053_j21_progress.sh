# TIMEOUT=120
# J21: no push for ~90 min after rel-f1/driver-dnf. Show where it is.
id=$(cat agent/state/J21.id); squeue -j "$id" -o "%.12i %.22j %.8T %.10M %.10l %R"
d=$(ls -td artifacts/runs/*J21_model_layers_* 2>/dev/null | head -1); echo "== $d"
wc -l < "$d/events.jsonl"; tail -2 "$d/events.jsonl" | cut -c1-260
echo "last event $(( $(date +%s) - $(stat -c %Y "$d/events.jsonl") ))s ago"
f=$(ls -t sbatch/logs/J21_model_layers_${id}*.out 2>/dev/null | head -1); [ -n "$f" ] && grep -a -E "^--- |FAILED|Error|Traceback" "$f" | tail -5 | cut -c1-200
